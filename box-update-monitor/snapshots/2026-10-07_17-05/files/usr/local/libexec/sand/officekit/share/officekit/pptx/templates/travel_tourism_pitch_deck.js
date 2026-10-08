/*
 * Travel / Tourism pitch deck - 30 slides, 13.333 x 7.5 in (16:9).
 * Rebuilt from the reference .pptx with pptxgenjs only.
 * Photographs in the original are redrawn as flat grey placeholders.
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const RED       = 'B80005';
const RED_L     = 'F64F50';
const PINK      = 'FC7B7C';
const CORAL     = 'FD6665';
const RED_B     = 'EE2122';
const RED_D     = 'D12828';
const BLACK     = '000000';
const WHITE     = 'FFFFFF';
const CHARCOAL  = '262626';
const GRAY      = '595959';
const SNOW      = 'F2F2F2';
const PHOTO     = 'A6A6A6';

const HEAD = 'Bebas Neue';   // theme major font (condensed display)
const BODY = 'Open Sans';    // theme minor font

// ------------------------------------------------- repeated body copy
const T_LOREP_IPSUM_DUIS =
  'Lorep  ipsum duis aute irure dolor in kauselih oilusioi reprehenderiti volui ptates esse cill inure dolorlasue boru sit amet. Duis aute';
const T_THE_EUROPEAN_LANGUAGES =
  'The European languages are members of the same family. Their separate existence is a myth. ';
const T_LOREM_IPSUM_DOLOR =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Et dolor diam ultricies sed quisque. Tortor cursus sed blandit..';
const T_LOREP_IPSUM_DUIS2 =
  'Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau';
const T_LOREP_IPSUM_DUIS3 =
  'Lorep  ipsum duis aute irure dolor in kauselih oilue eprehe kau deriti vols esse cill inure dolorlaboru sit amet. Duis autelo jiuh irusitakus reprehenderi Voluptate lorem kuisais louisin kaesiul';
const T_DOLOR_IN_KAUSELIH =
  'dolor in kauselih oilue reprehend';
const T_DERITI_VOLS_ESSE =
  'deriti vols esse cill inure dolorlaboru sit amet. Duis autelo jiuh irusitakus reprehenderi Voluptate lorem kuisais louisin kaesiul';
const T_LOREM_IPSUM_DOLOR2 =
  'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ';
const T_WE_CREATE_EXPERIENCES =
  'We create experiences FOR TOURIST';
const T_UT_WISI_ENIM =
  'Ut wisi enim ad minim veniam, quis nostrud exerci';
const T_LOREM_IPSUM_DECIDED =
  'Lorem Ipsum decided to leave for the far World of Grammar. ';
const T_LOREM_IPSUM_DOLOR3 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam non feugiat mauris. In massa magna, aliquet sed feugiat eget, dignissim quis magna. Fusce ut nisi imperdiet, tempor leo ac, molestie magna. Nulla faucibus urna sit amet urna faucibus lobortis. ';
const T_A_TEXT_IS =
  'A text is easier to understand with a title';
const T_WE_CREATE_EXPERIENCES2 =
  'We create experiences that elevate brands.';
const T_SUSPENDISSE_DIAM_FELIS =
  'Suspendisse diam felis, tempus non nunc et, egestas blandit ex. Quisque in ante ipsum. Curabitur vehicula eget orci a pretium. Cras sapien nisi, dapibus at feugiat ac, faucibus ac arcu. Suspendisse augue arcu, accumsan eget bibendum vitae, blandit quis augue. ';
const T_ESSE_CILL_INURE =
  'esse cill inure dolorlaboru sit amet. Duis aute';

// ------------------------------- reusable outlines (unit space 0..1)
const BLOB = [
    ['M', 0, 0.5], ['C', 0, 0.125, 0.125, 0, 0.5, 0], ['C', 0.875, 0, 1, 0.125, 1, 0.5],
    ['C', 1, 0.875, 0.875, 1, 0.5, 1], ['C', 0.125, 1, 0, 0.875, 0, 0.5], ['Z']
  ];
const ARROW_TIP = [
    ['M', 1, 0.538], ['C', 0.227, 0.974, 0.227, 0.974, 0.227, 0.974], ['C', 0.182, 1, 0.182, 1, 0.136, 1],
    ['C', 0.091, 1, 0, 0.974, 0, 0.949], ['C', 0, 0.077, 0, 0.077, 0, 0.077],
    ['C', 0, 0.026, 0.091, 0, 0.136, 0], ['C', 0.182, 0, 0.182, 0.026, 0.227, 0.026],
    ['C', 1, 0.462, 1, 0.462, 1, 0.462], ['C', 1, 0.462, 1, 0.487, 1, 0.513],
    ['C', 1, 0.513, 1, 0.538, 1, 0.538], ['Z']
  ];
const CAPSULE = [
    ['M', 0, 0.5], ['C', 0, 0.172, 0.081, 0.035, 0.313, 0.006], ['L', 0.373, 0.003], ['L', 0.373, 0],
    ['L', 0.422, 0], ['L', 0.578, 0], ['L', 0.584, 0], ['L', 0.584, 0], ['L', 0.687, 0.006],
    ['C', 0.919, 0.035, 1, 0.172, 1, 0.5], ['C', 1, 0.828, 0.919, 0.965, 0.687, 0.994], ['L', 0.584, 1],
    ['L', 0.584, 1], ['L', 0.578, 1], ['L', 0.422, 1], ['L', 0.373, 1], ['L', 0.373, 0.997],
    ['L', 0.313, 0.994], ['C', 0.081, 0.965, 0, 0.828, 0, 0.5], ['Z']
  ];
const WEDGE_L = [['M', 0, 0], ['L', 0.944, 0], ['L', 1, 0], ['L', 1, 1], ['Z']];
const WIFI = [
    ['M', 0.447, 0.6], ['L', 0.447, 0.45], ['L', 0.553, 0.45], ['L', 0.553, 0.6], ['L', 0.816, 0.6],
    ['C', 0.845, 0.6, 0.869, 0.622, 0.869, 0.65], ['L', 0.869, 0.95],
    ['C', 0.869, 0.978, 0.845, 1, 0.816, 1], ['L', 0.184, 1], ['C', 0.155, 1, 0.131, 0.978, 0.131, 0.95],
    ['L', 0.131, 0.65], ['C', 0.131, 0.622, 0.155, 0.6, 0.184, 0.6], ['L', 0.447, 0.6], ['Z'],
    ['M', 0, 0.342], ['C', 0.07, 0.143, 0.267, 0, 0.5, 0], ['C', 0.733, 0, 0.93, 0.143, 1, 0.342],
    ['L', 0.9, 0.373], ['C', 0.826, 0.164, 0.588, 0.051, 0.367, 0.12],
    ['C', 0.241, 0.16, 0.142, 0.254, 0.1, 0.373], ['L', 0, 0.342], ['Z'], ['M', 0.2, 0.405],
    ['C', 0.255, 0.248, 0.434, 0.163, 0.6, 0.215], ['C', 0.694, 0.245, 0.769, 0.316, 0.8, 0.405],
    ['L', 0.7, 0.437], ['C', 0.663, 0.332, 0.544, 0.275, 0.433, 0.31],
    ['C', 0.37, 0.33, 0.321, 0.377, 0.3, 0.437], ['L', 0.2, 0.405], ['Z'], ['M', 0.237, 0.7],
    ['L', 0.237, 0.9], ['L', 0.763, 0.9], ['L', 0.763, 0.7], ['L', 0.237, 0.7], ['Z']
  ];
const ARROW_NE = [
    ['M', 1, 0], ['L', 0.034, 0], ['L', 0.034, 0.163], ['L', 0.719, 0.163], ['L', 0, 0.882],
    ['L', 0.118, 1], ['L', 0.837, 0.281], ['L', 0.837, 0.966], ['L', 1, 0.966], ['Z']
  ];

// ------------------------------------------------------------------ helpers
// Every drawing call takes a box [x, y, w, h] in inches plus plain options.

/**
 * Outlines PowerPoint derives from an adjustment handle. pptxgenjs can only
 * pass `rectRadius` / `angleRange`, so these three are rebuilt as real paths.
 * `a` is the adjustment as a fraction; `d` is its length along the short side.
 */
const PRESET_PATH = {
  // leaning box - top edge slid right by d
  parallelogram: (w, h, a) => {
    const d = a * Math.min(w, h) / w;
    return [['M', 0, 1], ['L', d, 0], ['L', 1, 0], ['L', 1 - d, 1], ['Z']];
  },
  // pentagon with a point on the right edge
  homePlate: (w, h, a) => {
    const d = a * Math.min(w, h) / w;
    return [['M', 0, 0], ['L', 1 - d, 0], ['L', 1, 0.5], ['L', 1 - d, 1], ['L', 0, 1], ['Z']];
  },
  // right arrow with a v-shaped notch cut into its tail
  notchedRightArrow: (w, h, a) => {
    const d = a * Math.min(w, h) / w, y = 0.5 - a / 2;
    return [['M', 0, y], ['L', 1 - d, y], ['L', 1 - d, 0], ['L', 1, 0.5], ['L', 1 - d, 1],
            ['L', 1 - d, 1 - y], ['L', 0, 1 - y], ['L', d * (1 - 2 * y), 0.5], ['Z']];
  }
};

function scalePath(pts, w, h) {
  return pts.map(c => {
    if (c[0] === 'Z') return { close: true };
    if (c[0] === 'C') {
      return { x: c[5] * w, y: c[6] * h,
               curve: { type: 'cubic', x1: c[1] * w, y1: c[2] * h, x2: c[3] * w, y2: c[4] * h } };
    }
    return { x: c[1] * w, y: c[2] * h, moveTo: c[0] === 'M' };
  });
}

function boxOpts(b, o) {
  const t = { x: b[0], y: b[1], w: b[2], h: b[3] };
  if (o.fill) t.fill = typeof o.fill === 'string' ? { color: o.fill } : o.fill;
  if (o.line) t.line = o.line;
  if (o.radius !== undefined) t.rectRadius = o.radius * Math.min(b[2], b[3]);
  if (o.adj !== undefined) o.pts = PRESET_PATH[o.geom](b[2], b[3], o.adj);
  if (o.arc) { t.angleRange = [o.arc[0], o.arc[1]]; t.arcThicknessRatio = o.arc[2]; }
  if (o.shadow) t.shadow = Object.assign({ type: 'outer', rotateWithShape: 0 }, o.shadow);
  if (o.rot) t.rotate = o.rot;
  if (o.flipH) t.flipH = true;
  if (o.flipV) t.flipV = true;
  if (o.pts) t.points = scalePath(o.pts, b[2], b[3]);
  return t;
}

/** Plain shape: sh(slide, box, { geom, fill, line, radius, arc, pts, rot, shadow }) */
function sh(s, b, o) {
  const t = boxOpts(b, o || {});
  s.addShape(t.points ? 'custGeom' : ((o && o.geom) || 'rect'), t);
}

/** Straight rule drawn as a thin line shape. */
function rule(s, b, o) {
  s.addShape('line', Object.assign(boxOpts(b, o), { line: o.line }));
}

/**
 * Text block. `base` carries shape + default run styling, `paras` is one entry
 * per paragraph: a plain string, { t, ...runOpts } or { r: [runs], ...paraOpts }.
 */
function txt(s, b, base, paras) {
  base = base || {};
  const o = boxOpts(b, base);
  o.shape = o.points ? 'custGeom' : (base.geom || 'rect');
  o.valign = base.vc ? 'middle' : base.vb ? 'bottom' : 'top';
  o.align = base.align || 'left';
  o.fontFace = base.head ? HEAD : BODY;
  o.fontSize = base.sz || 18;
  o.color = base.col || WHITE;
  o.charSpacing = base.spc;
  o.bold = !!base.bold;
  o.italic = !!base.ital;
  o.underline = base.und ? { style: 'sng' } : undefined;
  o.lineSpacingMultiple = base.lh;
  o.paraSpaceBefore = base.before;
  o.paraSpaceAfter = base.after;
  if (base.nowrap) o.wrap = false;
  if (base.margin) o.margin = base.margin;
  if (!base.fill) o.fill = undefined;

  const runs = [];
  paras.forEach((p, pi) => {
    const para = typeof p === 'string' ? { t: p } : p;
    const parts = para.r || [para];
    parts.forEach((raw, ri) => {
      const r = typeof raw === 'string' ? { t: raw } : raw;
      runs.push({
        text: r.t,
        options: {
          fontSize: r.sz || o.fontSize,
          fontFace: r.head ? HEAD : (r.head === undefined && base.head ? HEAD : BODY),
          color: r.col || o.color,
          bold: r.bold !== undefined ? !!r.bold : o.bold,
          italic: r.ital !== undefined ? !!r.ital : o.italic,
          underline: r.und ? { style: 'sng' } : o.underline,
          charSpacing: r.spc || o.charSpacing,
          align: para.align || o.align,
          lineSpacingMultiple: para.lh || o.lineSpacingMultiple,
          paraSpaceBefore: para.before || o.paraSpaceBefore,
          paraSpaceAfter: para.after || o.paraSpaceAfter,
          breakLine: ri === parts.length - 1 && pi < paras.length - 1
        }
      });
    });
  });
  s.addText(runs, o);
}

/**
 * Device mock-ups redrawn as stacked rounded slabs. Each entry is
 * [x0, y0, x1, y1, cornerRadius, colour] as a fraction of the picture box;
 * the grey screen sits on top of these as a separate `photo` call.
 */
const DEVICE_SHELL = {
  phone:       [[0.055, 0.039, 0.947, 0.961, 0.10, '17161A']],
  watchSport:  [[0.165, 0.013, 0.835, 0.988, 0.10, 'EFEEEC'],
                [0.061, 0.240, 0.939, 0.760, 0.26, '9C9C9C']],
  watchAlpine: [[0.170, 0.011, 0.830, 0.991, 0.10, 'EAE1D8'],
                [0.058, 0.255, 0.977, 0.745, 0.26, 'BDB4AB']],
  tablet:      [[0.042, 0.033, 0.992, 0.967, 0.05, '141414']],
  laptop:      [[0.105, 0.065, 0.895, 0.905, 0.02, '0D0D10'],
                [0.015, 0.925, 0.985, 0.978, 0.50, '85868A']]
};

/** Stand-in for a bitmap: flat grey block clipped to `pts`, or a device shell. */
function photo(s, b, o) {
  o = o || {};
  if (o.device) {
    DEVICE_SHELL[o.device].forEach(p => {
      sh(s, [b[0] + p[0] * b[2], b[1] + p[1] * b[3], (p[2] - p[0]) * b[2], (p[3] - p[1]) * b[3]],
         { geom: 'roundRect', radius: p[4], fill: p[5] });
    });
    return;
  }
  sh(s, b, { pts: o.pts, fill: PHOTO });
  if (b[2] > 1.2 && b[3] > 0.8) {
    txt(s, [b[0], b[1] + b[3] / 2 - 0.16, b[2], 0.32],
        { sz: 10, col: '7C7C7C', align: 'center', vc: 1 }, ['[image]']);
  }
}

// ------------------------------------------------------------------------
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [0.957, 1.269, 11.42, 4.962], { geom: 'roundRect', radius: 0, fill: RED });
  photo(s, [0, 0, 13.333, 7.5]);
  txt(s, [11.076, 0.554, 1.904, 0.286], { sz: 11, head: 1, col: BLACK, align: 'right' }, ['MODERN DESIGN']);
  txt(s, [11.076, 6.66, 1.904, 0.286], { sz: 11, head: 1, col: BLACK, align: 'right' }, ['COVER']);
  txt(s, [0.313, 0.554, 1.904, 0.286], { sz: 11, head: 1, col: BLACK }, ['MODERN DESIGN']);
  txt(s, [0.313, 6.66, 1.904, 0.286], { sz: 11, head: 1, col: BLACK }, ['COVER']);
  txt(s, [0.957, 1.64, 11.42, 3.294], { sz: 115, bold: 1, head: 1, align: 'center', lh: 0.8 },
    ['TRAVEL', 'TOURISM TRAVEL']);
  txt(s, [3.515, 5.166, 6.303, 0.374], { spc: 3, align: 'center', lh: 0.9 }, ['INSERT YOUR TAGLINE HERE']);
}

// ------------------------------------------------------------------------
function slide02(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  photo(s, [7.484, 0, 5.849, 7.5]);
  txt(s, [1.027, 4.036, 4.42, 0.76], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS3]);
  txt(s, [0.984, 1.269, 4.463, 1.767], { sz: 33, head: 1 }, ['A MUST VISIT', 'EUROPE', 'EDITION']);
  txt(s, [1.027, 3.75, 3.478, 0.286], { sz: 11, head: 1 }, ['Modern Design']);
  sh(s, [1.115, 5.525, 0.516, 0.516], { geom: 'ellipse', fill: RED });
  txt(s, [1.122, 5.696, 0.5, 0.172], { sz: 10, col: BLACK, vc: 1, margin: [7.2, 7.2, 7.2, 7.2] }, ['+33']);
  txt(s, [1.755, 5.524, 1.557, 0.286], { sz: 11 }, ['YOUR TEXT HERE']);
  txt(s, [1.755, 5.735, 1.821, 0.283], { sz: 8, lh: 1.5 }, ['Lorem ipsum dolor ']);
  sh(s, [3.359, 5.525, 0.516, 0.516], { geom: 'ellipse', fill: RED_L });
  txt(s, [3.367, 5.696, 0.5, 0.172], { sz: 10, col: BLACK, vc: 1, margin: [7.2, 7.2, 7.2, 7.2] }, ['+23']);
  txt(s, [4, 5.524, 1.557, 0.286], { sz: 11 }, ['YOUR TEXT HERE']);
  txt(s, [4, 5.735, 1.821, 0.283], { sz: 8, lh: 1.5 }, ['Lorem ipsum dolor ']);
  sh(s, [1.122, 6.389, 4.323, 0], { geom: 'line', line: { color: WHITE } });
  sh(s, [1.151, 5.128, 4.323, 0], { geom: 'line', line: { color: WHITE } });
}

// ------------------------------------------------------------------------
function slide03(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  photo(s, [0, 0, 6.589, 6.191]);
  sh(s, [0.328, 3.691, 3.621, 2.5], { fill: RED });
  txt(s, [0.605, 4.138, 3.125, 0.672], { sz: 16, head: 1 }, ['22.000']);
  txt(s, [0.605, 4.895, 3.125, 0.881], { sz: 9, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR]);
  txt(s, [7.62, 0.647, 0.941, 0.286], { sz: 11, head: 1 }, ['Modern']);
  txt(s, [7.62, 0.968, 0.941, 0.286], { sz: 11, head: 1 }, ['Design']);
  txt(s, [7.62, 1.289, 0.941, 0.286], { sz: 11, head: 1 }, ['Pitch']);
  txt(s, [7.62, 3.51, 5.385, 1.582], { sz: 44, head: 1 }, [T_WE_CREATE_EXPERIENCES]);
  sh(s, [2.217, 6.401, 4.332, 0], { geom: 'line', line: { color: WHITE } });
  txt(s, [5.328, 5.777, 1.26, 0.413], { fill: RED, sz: 11, und: 1, align: 'center', vc: 1 }, ['Button Here']);
}

// ------------------------------------------------------------------------
function slide04(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  photo(s, [1.111, 0, 5.477, 7.5],
    { pts: [
    ['M', 0.05, 0], ['L', 0.95, 0], ['C', 0.978, 0, 1, 0.016, 1, 0.037], ['L', 1, 0.963],
    ['C', 1, 0.981, 0.983, 0.996, 0.96, 0.999], ['L', 0.95, 1], ['L', 0.05, 1], ['L', 0.04, 0.999],
    ['C', 0.017, 0.996, 0, 0.981, 0, 0.963], ['L', 0, 0.037], ['C', 0, 0.016, 0.022, 0, 0.05, 0], ['Z']
  ] });
  txt(s, [1.707, 4.232, 4.42, 0.802], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS3]);
  txt(s, [1.664, 1.269, 4.463, 1.212], { sz: 33, head: 1 }, ['Your Gateway to EXPLORE THE WORLD']);
  txt(s, [1.707, 3.946, 3.478, 0.286], { sz: 11, head: 1 }, ['Modern Design']);
  txt(s, [1.795, 5.778, 2.245, 0.413],
    { geom: 'roundRect', radius: 0.5, fill: 'DC3E42', sz: 11.25, align: 'center', vc: 1 }, ['Button Here']);
  txt(s, [10.276, 1.641, 2.407, 0.802], { sz: 9, lh: 1.5 }, [T_THE_EUROPEAN_LANGUAGES]);
  txt(s, [7.185, 1.657, 1.314, 0.286], { flipH: 1, sz: 11, bold: 1, head: 1 }, ['Modern Design']);
  sh(s, [8.643, 1.859, 1.154, 0], { geom: 'line', line: { color: RED, width: 1 }, flipH: 1 });
  txt(s, [10.276, 3.495, 2.407, 0.802], { sz: 9, lh: 1.5 }, [T_THE_EUROPEAN_LANGUAGES]);
  txt(s, [7.185, 3.512, 1.314, 0.286], { flipH: 1, sz: 11, bold: 1, head: 1 }, ['Modern Design']);
  sh(s, [8.643, 3.714, 1.154, 0], { geom: 'line', line: { color: RED, width: 1 }, flipH: 1 });
  txt(s, [10.276, 5.35, 2.407, 0.802], { sz: 9, lh: 1.5 }, [T_THE_EUROPEAN_LANGUAGES]);
  txt(s, [7.185, 5.366, 1.314, 0.286], { flipH: 1, sz: 11, bold: 1, head: 1 }, ['Modern Design']);
  sh(s, [8.643, 5.568, 1.154, 0], { geom: 'line', line: { color: RED, width: 1 }, flipH: 1 });
}

// ------------------------------------------------------------------------
function slide05(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [8.94, -0.059, 4.393, 7.559], { pts: WEDGE_L, fill: RED, flipV: 1 });
  sh(s, [1.381, 1.826, 10.761, 2.289],
    { geom: 'parallelogram', adj: 0.25, fill: CHARCOAL, shadow: { blur: 100, offset: 0, angle: 0, color: BLACK, opacity: 0.21 } });
  txt(s, [2.509, 2.784, 8.829, 0.974], { sz: 12, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR3]);
  txt(s, [2.741, 2.187, 5.396, 0.438], { sz: 20, head: 1 }, [T_A_TEXT_IS]);
  sh(s, [2.509, 2.347, 0.068, 0.118], { pts: ARROW_TIP, fill: RED });
  txt(s, [2.863, 0.904, 7.608, 0.656], { sz: 33, head: 1, align: 'center' }, ['Two important tourist guide']);
  txt(s, [4.928, 0.568, 3.478, 0.337], { sz: 14, head: 1, col: BLACK, align: 'center' }, ['Subtitle Here']);
  sh(s, [1.381, 4.461, 10.761, 2.289],
    { geom: 'parallelogram', adj: 0.25, fill: CHARCOAL, shadow: { blur: 100, offset: 0, angle: 0, color: BLACK, opacity: 0.21 } });
  txt(s, [2.741, 4.873, 6.764, 0.438], { sz: 20, head: 1 }, [T_WE_CREATE_EXPERIENCES2]);
  sh(s, [2.509, 5.032, 0.068, 0.118], { pts: ARROW_TIP, fill: RED });
  txt(s, [2.509, 5.47, 8.64, 0.974], { sz: 12, lh: 1.5 }, [T_SUSPENDISSE_DIAM_FELIS]);
}

// ------------------------------------------------------------------------
function slide06(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [1.958, 4.17, 2.23, 1.924],
    { geom: 'parallelogram', adj: 0.25, fill: RED, shadow: { blur: 100, offset: 0, angle: 0, color: BLACK, opacity: 0.21 } });
  sh(s, [1.958, 1.534, 2.23, 1.924],
    { geom: 'parallelogram', adj: 0.25, fill: CHARCOAL, shadow: { blur: 100, offset: 0, angle: 0, color: BLACK, opacity: 0.21 } });
  txt(s, [4.884, 2.132, 6.491, 1.277], { sz: 12, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR3]);
  txt(s, [5.116, 1.534, 5.548, 0.438], { sz: 20, head: 1 }, [T_A_TEXT_IS]);
  sh(s, [4.884, 1.694, 0.068, 0.118], { pts: ARROW_TIP, fill: RED });
  txt(s, [5.116, 4.22, 6.352, 0.438], { sz: 20, head: 1 }, [T_WE_CREATE_EXPERIENCES2]);
  sh(s, [4.884, 4.38, 0.068, 0.118], { pts: ARROW_TIP, fill: RED });
  txt(s, [4.884, 4.817, 6.352, 1.277], { sz: 12, lh: 1.5 }, [T_SUSPENDISSE_DIAM_FELIS]);
  sh(s, [2.789, 2.161, 0.569, 0.671],
    { pts: [
    ['M', 0.123, 0.676], ['C', 0.111, 0.63, 0.105, 0.583, 0.105, 0.536],
    ['C', 0.105, 0.292, 0.269, 0.084, 0.5, 0], ['C', 0.731, 0.084, 0.895, 0.292, 0.895, 0.536],
    ['C', 0.895, 0.584, 0.889, 0.631, 0.876, 0.676], ['L', 0.991, 0.762],
    ['C', 1.001, 0.77, 1.003, 0.782, 0.996, 0.792], ['L', 0.857, 0.988],
    ['C', 0.849, 1, 0.832, 1.003, 0.819, 0.997], ['C', 0.817, 0.996, 0.815, 0.994, 0.813, 0.993],
    ['L', 0.686, 0.885], ['C', 0.675, 0.876, 0.661, 0.871, 0.646, 0.871], ['L', 0.354, 0.871],
    ['C', 0.339, 0.871, 0.325, 0.876, 0.314, 0.885], ['L', 0.187, 0.993],
    ['C', 0.176, 1.002, 0.158, 1.002, 0.147, 0.993], ['C', 0.145, 0.992, 0.144, 0.99, 0.143, 0.988],
    ['L', 0.004, 0.792], ['C', -0.003, 0.782, -0.001, 0.77, 0.009, 0.762], ['L', 0.123, 0.676], ['Z'],
    ['M', 0.5, 0.536], ['C', 0.562, 0.536, 0.613, 0.493, 0.613, 0.44],
    ['C', 0.613, 0.387, 0.562, 0.344, 0.5, 0.344], ['C', 0.438, 0.344, 0.387, 0.387, 0.387, 0.44],
    ['C', 0.387, 0.493, 0.438, 0.536, 0.5, 0.536], ['Z']
  ], fill: RED });
  sh(s, [2.752, 4.861, 0.643, 0.643],
    { pts: [
    ['M', 0.5, 1], ['C', 0.224, 1, 0, 0.776, 0, 0.5], ['C', 0, 0.224, 0.224, 0, 0.5, 0],
    ['C', 0.776, 0, 1, 0.224, 1, 0.5], ['C', 1, 0.776, 0.776, 1, 0.5, 1], ['Z'], ['M', 0.725, 0.275],
    ['L', 0.4, 0.4], ['L', 0.275, 0.725], ['L', 0.6, 0.6], ['L', 0.725, 0.275], ['Z'], ['M', 0.5, 0.55],
    ['C', 0.472, 0.55, 0.45, 0.528, 0.45, 0.5], ['C', 0.45, 0.472, 0.472, 0.45, 0.5, 0.45],
    ['C', 0.528, 0.45, 0.55, 0.472, 0.55, 0.5], ['C', 0.55, 0.528, 0.528, 0.55, 0.5, 0.55], ['Z']
  ], fill: BLACK });
}

// ------------------------------------------------------------------------
function slide07(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [4.138, 2.651, 8.116, 1.96], { geom: 'parallelogram', adj: 0.25, fill: RED });
  txt(s, [8.108, 1.296, 3.262, 1.031], { sz: 12, lh: 1.5 }, [T_THE_EUROPEAN_LANGUAGES]);
  txt(s, [5.661, 1.491, 1.92, 0.337], { flipH: 1, sz: 14, bold: 1, head: 1 }, ['Modern Design']);
  sh(s, [7.356, 1.693, 0.488, 0], { geom: 'line', line: { color: RED, width: 1 }, flipH: 1 });
  txt(s, [8.108, 3.15, 3.262, 1.031], { sz: 12, lh: 1.5 }, [T_THE_EUROPEAN_LANGUAGES]);
  txt(s, [5.661, 3.345, 1.92, 0.337], { flipH: 1, sz: 14, bold: 1, head: 1 }, ['Modern Design']);
  sh(s, [7.356, 3.547, 0.488, 0], { geom: 'line', line: { color: WHITE, width: 1 }, flipH: 1 });
  txt(s, [8.108, 5.005, 3.262, 1.031], { sz: 12, lh: 1.5 }, [T_THE_EUROPEAN_LANGUAGES]);
  txt(s, [5.661, 5.2, 1.92, 0.337], { flipH: 1, sz: 14, bold: 1, head: 1 }, ['Modern Design']);
  sh(s, [7.356, 5.402, 0.488, 0], { geom: 'line', line: { color: RED, width: 1 }, flipH: 1 });
  txt(s, [1.08, 2.734, 3.058, 1.447], { sz: 40, head: 1 }, ['Travel', 'point']);
}

// ------------------------------------------------------------------------
function slide08(pptx) {
  const s = pptx.addSlide();
  s.background = { color: RED };
  photo(s, [0, 0.443, 2.688, 7.057], { pts: [['M', 0, 0], ['L', 1, 1], ['L', 0.056, 1], ['L', 0, 1], ['Z']] });
  txt(s, [2.069, 3.151, 9.194, 1.995], { sz: 12, align: 'center', lh: 1.5 },
    [{ t: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam non feugiat mauris. In massa magna, aliquet sed feugiat eget, dignissim quis magna. Fusce ut nisi imperdiet, tempor leo ac, molestie magna. Nulla faucibus urna sit amet urna faucibus lobortis. Suspendisse diam felis, tempus non nunc', after: 7.5 }, { t: 'Duis condimentum orci rutrum, ornare risus non, iaculis risus. Aenean vestibulum nisi felis, vitae ullamcorper mi tristique et. Nulla nulla massa, congue gravida pulvinar in, ornare non ligula. Praesent eu felis magna. Proin sem odio, auctor nec rhoncus vel, laoreet vitae tellus. Donec suscipit porttitor facilisis.', after: 7.5 }]);
  txt(s, [5.193, 0.802, 2.947, 0.337], { sz: 14, head: 1, align: 'center' }, ['Single paragraph content']);
  txt(s, [2.549, 5.802, 2.558, 0.567],
    { geom: 'roundRect', radius: 0.5, fill: BLACK, line: { color: BLACK }, sz: 11.25, align: 'center', vc: 1 },
    ['Option One']);
  txt(s, [5.388, 5.802, 2.558, 0.567],
    { geom: 'roundRect', radius: 0.5, fill: BLACK, line: { color: BLACK }, sz: 11.25, align: 'center', vc: 1 },
    ['Option Two']);
  txt(s, [8.227, 5.802, 2.558, 0.567],
    { geom: 'roundRect', radius: 0.5, fill: BLACK, line: { color: BLACK }, sz: 11.25, align: 'center', vc: 1 },
    ['Option Three']);
  txt(s, [3.076, 1.131, 7.182, 1.582], { sz: 44, head: 1, align: 'center' }, ['Your Gateway to ', 'EXPLORE THE WORLD']);
  sh(s, [0.95, 3.478, 1.844, 0.494], { geom: 'parallelogram', adj: 0.864, fill: RED_L, flipH: 1 });
  sh(s, [10.559, -0.003, 1.844, 0.494], { geom: 'parallelogram', adj: 0.864, fill: RED_L, flipH: 1 });
  photo(s, [10.646, 0, 2.688, 7.057], { pts: WEDGE_L });
}

// ------------------------------------------------------------------------
function slide09(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  photo(s, [0, 0, 13.333, 3.75]);
  sh(s, [0, 3.75, 4.444, 3.75], { fill: BLACK });
  sh(s, [8.889, 3.75, 4.444, 3.75], { fill: BLACK });
  sh(s, [0, 0, 4.444, 3.75], { fill: RED });
  sh(s, [4.444, 0, 1.333, 3.75], { fill: { color: RED, transparency: 36 } });
  sh(s, [4.876, 2.81, 0.47, 0.47],
    { geom: 'ellipse', line: { color: BLACK, width: 2 }, shadow: { blur: 40, offset: 35, angle: 45, color: BLACK, opacity: 0.1 } });
  sh(s, [5.031, 2.965, 0.16, 0.16],
    { pts: [
    ['M', 0.221, 0], ['L', 1, 0], ['L', 1, 0.779], ['L', 0.779, 1], ['L', 0.779, 0.376], ['L', 0.156, 1],
    ['L', 0, 0.844], ['L', 0.624, 0.221], ['L', 0, 0.221], ['Z']
  ], fill: BLACK });
  txt(s, [-0.078, 1.069, 3.956, 1.919], { sz: 36, head: 1, align: 'right' }, [T_WE_CREATE_EXPERIENCES]);
  txt(s, [0.399, 0.733, 3.478, 0.337], { sz: 14, head: 1, align: 'right' }, ['Subtitle Here']);
  txt(s, [2.764, 6.582, 0.784, 0.241],
    { geom: 'roundRect', radius: 0.5, fill: RED, sz: 11, head: 1, align: 'center', vc: 1 }, ['2024']);
  txt(s, [0.728, 4.405, 2.82, 0.39], { sz: 16, bold: 1, head: 1, lh: 1.1 }, ['Lorep  ipsum duis aute']);
  txt(s, [0.728, 4.784, 2.82, 0.602], { sz: 12, lh: 1.3 }, [T_UT_WISI_ENIM]);
  txt(s, [7.293, 6.582, 0.784, 0.241],
    { geom: 'roundRect', radius: 0.5, fill: RED, sz: 11, head: 1, align: 'center', vc: 1 }, ['2024']);
  txt(s, [5.257, 4.405, 2.82, 0.39], { sz: 16, bold: 1, head: 1, lh: 1.1 }, ['Lorep  ipsum duis aute']);
  txt(s, [5.257, 4.784, 2.82, 0.602], { sz: 12, lh: 1.3 }, [T_UT_WISI_ENIM]);
  txt(s, [11.737, 6.582, 0.784, 0.241],
    { geom: 'roundRect', radius: 0.5, fill: RED, sz: 11, head: 1, align: 'center', vc: 1 }, ['2024']);
  txt(s, [9.701, 4.405, 2.82, 0.39], { sz: 16, bold: 1, head: 1, lh: 1.1 }, ['Lorep  ipsum duis aute']);
  txt(s, [9.701, 4.784, 2.82, 0.602], { sz: 12, lh: 1.3 }, [T_UT_WISI_ENIM]);
}

// ------------------------------------------------------------------------
function slide10(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  photo(s, [6.011, 0, 7.322, 7.5],
    { pts: [['M', 0.594, 0], ['L', 0.784, 0], ['L', 0.81, 0], ['L', 1, 0], ['L', 1, 1], ['L', 0, 1], ['Z']] });
  photo(s, [3.089, 0, 6.995, 6.03], { pts: [['M', 0, 0], ['L', 1, 0], ['L', 0.5, 1], ['Z']] });
  sh(s, [3.558, 3.75, 4.372, 2.443],
    { geom: 'parallelogram', adj: 0.25, fill: CHARCOAL, shadow: { blur: 100, offset: 0, angle: 0, color: BLACK, opacity: 0.21 } });
  txt(s, [3.978, 5.195, 2.947, 0.66], { sz: 11, lh: 1.5 },
    [{ r: ['Lorep  ipsum duis ', { t: 'aute irure ', bold: 1 }, T_DOLOR_IN_KAUSELIH] }]);
  txt(s, [4.364, 4.063, 2.017, 0.404], { head: 1 }, ['Presentation Design']);
  sh(s, [1.391, 1.059, 4.372, 2.443],
    { geom: 'parallelogram', adj: 0.25, fill: RED, shadow: { blur: 100, offset: 0, angle: 0, color: BLACK, opacity: 0.21 } });
  txt(s, [1.807, 2.454, 3.395, 0.66], { sz: 11, lh: 1.5 },
    [{ r: ['Lorep  ipsum duis ', { t: 'aute irure ', bold: 1 }, T_DOLOR_IN_KAUSELIH] }]);
  txt(s, [2.232, 1.322, 2.017, 0.404], { head: 1 }, ['Presentation Design']);
}

// ------------------------------------------------------------------------
function slide11(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [0, 3.476, 9.477, 3.582],
    { pts: [['M', 0, 0], ['L', 1, 0], ['L', 0.849, 1], ['L', 0, 1], ['Z']], fill: BLACK });
  txt(s, [1.391, 3.948, 4.378, 1.178], { sz: 32, head: 1 }, [T_WE_CREATE_EXPERIENCES]);
  txt(s, [1.391, 5.875, 4.378, 0.71], { sz: 12, lh: 1.5 },
    [{ r: ['Lorep  ipsum duis ', { t: 'aute irure ', bold: 1 }, 'dolor in kauselih oilue reprehend esse cill Voluptate lorem kuisais.'] }]);
  txt(s, [1.391, 5.437, 2.947, 0.337], { sz: 14, head: 1 }, ['Presentation Design']);
  photo(s, [8.204, 0.443, 7.823, 6.615], { pts: [['M', 0.332, 0], ['L', 1, 0], ['L', 0.668, 1], ['L', 0, 1], ['Z']] });
  photo(s, [-1.208, 0.443, 11.884, 3.034],
    { pts: [['M', 0.102, 0], ['L', 1, 0], ['L', 0.898, 1], ['L', 0, 1], ['Z']] });
}

// ------------------------------------------------------------------------
function slide12(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [0, 0, 7.485, 7.5], { geom: 'roundRect', radius: 0.1, fill: BLACK });
  photo(s, [1.111, 0, 12.222, 7.5]);
  txt(s, [11.076, 0.554, 1.904, 0.286], { sz: 11, head: 1, align: 'right' }, ['Modern Design']);
  txt(s, [1.564, 0.431, 2.901, 0.562], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS2]);
  sh(s, [7.485, 4.578, 5.446, 2.541], { geom: 'roundRect', radius: 0.1, fill: CHARCOAL });
  txt(s, [7.838, 5.918, 4.332, 0.723], { sz: 8, lh: 1.5 },
    [{ r: ['Lorep  ipsum duis ', { t: 'aute irure ', bold: 1 }, T_DOLOR_IN_KAUSELIH] }, { r: [T_ESSE_CILL_INURE, { t: ' irusitakus ', bold: 1 }, 'reprei'] }, 'Voluptate lorem kuisais.']);
  txt(s, [7.838, 4.885, 2.947, 0.337], { sz: 14, head: 1 }, ['Presentation Design']);
  txt(s, [10.909, 4.865, 1.26, 0.413],
    { geom: 'roundRect', radius: 0.5, fill: RED, sz: 11.25, align: 'center', vc: 1 }, ['Button Here']);
  txt(s, [1.673, 1.772, 7.391, 1.582], { sz: 44, head: 1 }, ['Tourism', 'edition']);
}

// ------------------------------------------------------------------------
function slide13(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [5.609, 4.382, 7.72, 0.549], { fill: RED });
  txt(s, [5.814, 4.522, 0.32, 0.267], { sz: 10, bold: 1, lh: 1.2, nowrap: 1, margin: [0, 11.3, 2.8, 2.8] },
    [{ t: '02', before: 10 }]);
  txt(s, [9.158, 4.522, 0.32, 0.267], { sz: 10, bold: 1, lh: 1.2, nowrap: 1, margin: [0, 11.3, 2.8, 2.8] },
    [{ t: '03', before: 10 }]);
  txt(s, [2.501, 4.522, 0.32, 0.267], { sz: 10, bold: 1, lh: 1.2, nowrap: 1, margin: [0, 11.3, 2.8, 2.8] },
    [{ t: '01', before: 10 }]);
  txt(s, [5.842, 0.817, 5.825, 1.313], { sz: 72, head: 1 }, ['Agenda']);
  txt(s, [5.766, 2.692, 1.593, 0.286], { sz: 11 }, ['TEXT TITTLE HERE']);
  txt(s, [5.766, 2.978, 2.644, 0.986], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS]);
  txt(s, [8.996, 2.692, 1.593, 0.286], { sz: 11 }, ['TEXT TITTLE HERE']);
  txt(s, [8.996, 2.978, 2.644, 0.986], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS]);
  txt(s, [2.231, 5.128, 1.593, 0.286], { sz: 11 }, ['TEXT TITTLE HERE']);
  txt(s, [2.231, 5.415, 2.644, 0.986], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS]);
  photo(s, [2.224, 0.325, 3.222, 4.606]);
  photo(s, [5.609, 5.128, 3.222, 2.372]);
  photo(s, [8.997, 5.128, 3.222, 2.372]);
}

// ------------------------------------------------------------------------
function slide14(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [0, 0, 13.333, 3.55], { fill: RED });
  txt(s, [2.388, 0.529, 8.558, 0.656], { sz: 33, head: 1, align: 'center' }, ['Executive summary']);
  txt(s, [2.664, 1.411, 8.006, 0.559], { sz: 9, align: 'center', lh: 1.5 }, [T_LOREP_IPSUM_DUIS2, T_DERITI_VOLS_ESSE]);
  sh(s, [4.462, 2.444, 7.915, 4.072], { geom: 'roundRect', radius: 0.1, fill: CHARCOAL });
  txt(s, [8.85, 3.693, 3.435, 1.043], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS3]);
  txt(s, [8.85, 3.406, 3.478, 0.286], { sz: 11, head: 1 }, ['Modern Design']);
  txt(s, [8.938, 5.238, 2.245, 0.413],
    { geom: 'roundRect', radius: 0.5, fill: 'DC3E42', sz: 11.25, align: 'center', vc: 1 }, ['Button Here']);
  sh(s, [0.731, 2.444, 3.467, 4.072], { geom: 'roundRect', radius: 0.1, fill: CHARCOAL });
  txt(s, [1.114, 5.054, 2.211, 0.286], { sz: 11 }, ['Your Text Here']);
  txt(s, [1.114, 5.379, 3.003, 0.759], { sz: 9, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR]);
  sh(s, [1.178, 4.847, 0.337, 0.017], { fill: RED_L, line: { color: RED } });
  txt(s, [1.69, 3.446, 0.829, 0.656], { sz: 33, head: 1, align: 'center' }, ['01']);
  sh(s, [1.14, 2.781, 1.929, 1.941], { geom: 'blockArc', arc: [270.489, 66.696, 0.302], fill: RED });
  sh(s, [1.14, 2.779, 1.929, 1.941], { geom: 'blockArc', arc: [66.678, 148.958, 0.301], fill: RED_L });
  sh(s, [1.14, 2.779, 1.929, 1.941], { geom: 'blockArc', arc: [148.666, 183.95, 0.304], fill: PINK });
  sh(s, [1.14, 2.778, 1.929, 1.941], { geom: 'blockArc', arc: [182.214, 213.635, 0.303], fill: 'FEC2C1' });
  sh(s, [1.14, 2.78, 1.929, 1.941], { geom: 'blockArc', arc: [213.437, 270.909, 0.301], fill: RED_B });
  sh(s, [5.211, 2.881, 3.188, 3.208], { geom: 'blockArc', arc: [270.489, 66.696, 0.302], fill: RED });
  sh(s, [5.211, 2.879, 3.188, 3.208], { geom: 'blockArc', arc: [66.678, 148.958, 0.301], fill: RED_L });
  sh(s, [5.211, 2.879, 3.188, 3.208], { geom: 'blockArc', arc: [148.666, 182.278, 0.303], fill: PINK });
  sh(s, [5.211, 2.876, 3.188, 3.208], { geom: 'blockArc', arc: [182.214, 213.635, 0.303], fill: 'FEC2C1' });
  sh(s, [5.211, 2.88, 3.188, 3.208], { geom: 'blockArc', arc: [213.437, 270.909, 0.301], fill: RED_B });
  txt(s, [7.939, 4.303, 0.436, 0.286], { sz: 11, col: SNOW }, ['6.1']);
  txt(s, [6.166, 5.616, 0.469, 0.286], { sz: 11, col: SNOW }, ['3.2']);
  txt(s, [5.278, 4.667, 0.436, 0.286], { sz: 11, col: SNOW }, ['1.3']);
  txt(s, [5.294, 3.923, 0.436, 0.286], { sz: 11, col: SNOW }, ['1.2']);
  txt(s, [5.952, 3.187, 0.428, 0.286], { sz: 11, col: SNOW }, ['2.2']);
}

// ------------------------------------------------------------------------
function slide15(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [8.99, 0, 4.344, 7.5], { fill: CHARCOAL });
  sh(s, [8.849, 2.569, 4.484, 2.361], { fill: RED });
  txt(s, [0.984, 1.24, 8.558, 0.656], { sz: 33, head: 1 }, ['Executive summary']);
  txt(s, [0.984, 2.112, 5.761, 0.759], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS2, T_DERITI_VOLS_ESSE]);
  txt(s, [9.891, 3.365, 2.211, 0.286], { sz: 11 }, ['Your Text Here']);
  txt(s, [9.891, 3.691, 3.003, 0.759], { sz: 9, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR]);
  sh(s, [9.955, 3.158, 0.337, 0.017], { fill: RED_L, line: { color: RED } });
  txt(s, [9.891, 5.455, 2.211, 0.286], { sz: 11 }, ['Your Text Here']);
  txt(s, [9.891, 5.781, 3.003, 0.759], { sz: 9, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR]);
  sh(s, [9.955, 5.248, 0.337, 0.017], { fill: RED_L, line: { color: RED } });
  txt(s, [9.891, 1.319, 2.211, 0.286], { sz: 11 }, ['Your Text Here']);
  txt(s, [9.891, 1.644, 3.003, 0.759], { sz: 9, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR]);
  sh(s, [9.955, 1.111, 0.337, 0.017], { fill: RED_L, line: { color: RED } });
  txt(s, [0.984, 3.652, 0.601, 0.36], { sz: 12, nowrap: 1 }, ['100']);
  txt(s, [1.094, 4.16, 0.489, 0.36], { sz: 12, nowrap: 1 }, ['80']);
  txt(s, [1.096, 4.664, 0.489, 0.36], { sz: 12, nowrap: 1 }, ['60']);
  txt(s, [1.096, 5.172, 0.489, 0.36], { sz: 12, nowrap: 1 }, ['40']);
  txt(s, [1.097, 5.673, 0.489, 0.36], { sz: 12, nowrap: 1 }, ['20']);
  txt(s, [1.205, 6.18, 0.377, 0.36], { sz: 12, nowrap: 1 }, ['0']);
  sh(s, [2.069, 4.595, 0.665, 1.043], { geom: 'line', line: { color: RED_L, width: 1.5 }, flipV: 1 });
  sh(s, [2.72, 4.595, 0.647, 0.371], { geom: 'line', line: { color: RED_L, width: 1.5 } });
  sh(s, [3.367, 4.867, 0.651, 0.099], { geom: 'line', line: { color: RED_L, width: 1.5 }, flipV: 1 });
  sh(s, [4.018, 4.867, 0.652, 0.343], { geom: 'line', line: { color: RED_L, width: 1.5 } });
  sh(s, [4.67, 5.21, 0.666, 0.167], { geom: 'line', line: { color: RED_L, width: 1.5 } });
  sh(s, [5.321, 4.199, 0.659, 1.181], { geom: 'line', line: { color: RED_L, width: 1.5 }, flipV: 1 });
  sh(s, [5.961, 4.193, 0.643, 1.401], { geom: 'line', line: { color: RED_L, width: 1.5 } });
  sh(s, [6.603, 5.593, 0.68, 0.438], { geom: 'line', line: { color: RED_L, width: 1.5 } });
  sh(s, [7.281, 4.966, 0.668, 1.065], { geom: 'line', line: { color: RED_L, width: 1.5 }, flipV: 1 });
  sh(s, [7.898, 4.92, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [7.231, 5.967, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [6.547, 5.547, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [5.915, 4.17, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [5.266, 5.33, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [4.632, 5.166, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [3.97, 4.829, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [3.312, 4.911, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [2.678, 4.567, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [2.024, 5.574, 0.102, 0.093], { geom: 'ellipse', fill: RED_L });
  sh(s, [4.007, 4.178, 0.653, 0], { geom: 'line', line: { color: '8A0004', width: 1.5 } });
  sh(s, [2.05, 4.951, 0.665, 0.581], { geom: 'line', line: { color: '8A0004', width: 1.5 } });
  sh(s, [2.715, 4.852, 0.653, 0.68], { geom: 'line', line: { color: '8A0004', width: 1.5 }, flipV: 1 });
  sh(s, [3.366, 4.178, 0.642, 0.674], { geom: 'line', line: { color: '8A0004', width: 1.5 }, flipV: 1 });
  sh(s, [4.65, 4.178, 0.666, 0.167], { geom: 'line', line: { color: '8A0004', width: 1.5 } });
  sh(s, [5.311, 4.345, 0.643, 1.613], { geom: 'line', line: { color: '8A0004', width: 1.5 } });
  sh(s, [5.951, 5.568, 0.652, 0.393], { geom: 'line', line: { color: '8A0004', width: 1.5 }, flipV: 1 });
  sh(s, [6.6, 5.568, 0.648, 0.321], { geom: 'line', line: { color: '8A0004', width: 1.5 } });
  sh(s, [7.249, 5.811, 0.68, 0.081], { geom: 'line', line: { color: '8A0004', width: 1.5 }, flipV: 1 });
  sh(s, [2.018, 4.902, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [2.667, 5.475, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [3.953, 4.138, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [4.594, 4.138, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [5.26, 4.303, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [5.915, 5.911, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [6.553, 5.532, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [7.198, 5.851, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [7.868, 5.764, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [3.319, 4.806, 0.102, 0.093], { geom: 'ellipse', fill: '8A0004' });
  sh(s, [2.061, 5.298, 0.665, 0.449], { geom: 'line', line: { color: PINK, width: 1.5 } });
  sh(s, [2.726, 4.309, 0.64, 1.444], { geom: 'line', line: { color: PINK, width: 1.5 }, flipV: 1 });
  sh(s, [3.368, 4.311, 0.64, 1.521], { geom: 'line', line: { color: PINK, width: 1.5 } });
  sh(s, [3.998, 4.897, 0.652, 0.932], { geom: 'line', line: { color: PINK, width: 1.5 }, flipV: 1 });
  sh(s, [4.651, 4.897, 0.66, 0.401], { geom: 'line', line: { color: PINK, width: 1.5 } });
  sh(s, [5.31, 5.29, 0.642, 0.212], { geom: 'line', line: { color: PINK, width: 1.5 } });
  sh(s, [5.952, 5.502, 0.654, 0.095], { geom: 'line', line: { color: PINK, width: 1.5 } });
  sh(s, [6.592, 5.597, 0.675, 0.389], { geom: 'line', line: { color: PINK, width: 1.5 } });
  sh(s, [7.265, 4.534, 0.656, 1.453], { geom: 'line', line: { color: PINK, width: 1.5 }, flipV: 1 });
  sh(s, [2.032, 5.263, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
  sh(s, [2.668, 5.685, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
  sh(s, [3.313, 4.271, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
  sh(s, [3.947, 5.758, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
  sh(s, [4.607, 4.858, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
  sh(s, [5.257, 5.249, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
  sh(s, [5.884, 5.456, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
  sh(s, [6.539, 5.557, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
  sh(s, [7.214, 5.927, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
  sh(s, [7.86, 4.506, 0.102, 0.093], { geom: 'ellipse', fill: PINK });
}

// ------------------------------------------------------------------------
function slide16(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [0, 0, 13.333, 3.55], { fill: RED });
  txt(s, [2.388, 0.529, 8.558, 0.656], { sz: 33, head: 1, align: 'center' }, ['Executive summary']);
  txt(s, [2.664, 1.411, 8.006, 0.559], { sz: 9, align: 'center', lh: 1.5 }, [T_LOREP_IPSUM_DUIS2, T_DERITI_VOLS_ESSE]);
  sh(s, [1.136, 2.832, 5.43, 3.622], { geom: 'roundRect', radius: 0.1, fill: CHARCOAL });
  sh(s, [6.745, 2.832, 5.43, 3.622], { geom: 'roundRect', radius: 0.1, fill: CHARCOAL });
  sh(s, [2.459, 3.986, 0.504, 0.681], { fill: '59AAF2' });
  sh(s, [2.459, 4.667, 0.504, 1.176], { fill: '0F6FC6' });
  txt(s, [1.682, 5.746, 0.29, 0.286], { sz: 11, col: GRAY, align: 'right', nowrap: 1 }, ['0']);
  txt(s, [1.594, 5.269, 0.377, 0.286], { sz: 11, col: GRAY, align: 'right', nowrap: 1 }, ['20']);
  txt(s, [1.594, 4.797, 0.377, 0.286], { sz: 11, col: GRAY, align: 'right', nowrap: 1 }, ['40']);
  txt(s, [1.594, 4.294, 0.377, 0.286], { sz: 11, col: GRAY, align: 'right', nowrap: 1 }, ['60']);
  txt(s, [1.594, 3.804, 0.377, 0.286], { sz: 11, col: GRAY, align: 'right', nowrap: 1 }, ['80']);
  txt(s, [1.507, 3.344, 0.465, 0.286], { sz: 11, col: GRAY, align: 'right', nowrap: 1 }, ['100']);
  sh(s, [3.543, 3.951, 0.504, 1.211], { fill: '59AAF2' });
  sh(s, [3.543, 5.155, 0.504, 0.688], { fill: '0F6FC6' });
  sh(s, [4.615, 4.336, 0.504, 0.484], { fill: '59AAF2' });
  sh(s, [4.615, 4.82, 0.504, 1.023], { fill: '0F6FC6' });
  sh(s, [5.701, 3.727, 0.504, 1.205], { fill: '59AAF2' });
  sh(s, [5.699, 4.932, 0.504, 0.911], { fill: '0F6FC6' });
  txt(s, [7.16, 5.562, 0.298, 0.303], { sz: 12, col: GRAY, nowrap: 1 }, ['0']);
  txt(s, [7.16, 5.211, 0.298, 0.303], { sz: 12, col: GRAY, nowrap: 1 }, ['1']);
  txt(s, [7.16, 4.85, 0.298, 0.303], { sz: 12, col: GRAY, nowrap: 1 }, ['2']);
  txt(s, [7.16, 4.489, 0.298, 0.303], { sz: 12, col: GRAY, nowrap: 1 }, ['3']);
  txt(s, [7.16, 4.136, 0.298, 0.303], { sz: 12, col: GRAY, nowrap: 1 }, ['4']);
  txt(s, [7.16, 3.781, 0.298, 0.303], { sz: 12, col: GRAY, nowrap: 1 }, ['5']);
  txt(s, [7.16, 3.42, 0.298, 0.303], { sz: 12, col: GRAY, nowrap: 1 }, ['6']);
  sh(s, [8.011, 4.141, 1.146, 0.714], { geom: 'line', line: { color: RED_L, width: 2.25 }, flipV: 1 });
  sh(s, [9.145, 4.136, 1.443, 1.017], { geom: 'line', line: { color: RED_L, width: 2.25 } });
  sh(s, [10.582, 4.679, 1.178, 0.474], { geom: 'line', line: { color: RED_L, width: 2.25 }, flipV: 1 });
  sh(s, [8.023, 4.084, 1.146, 0.708], { geom: 'line', line: { color: RED, width: 2 } });
  sh(s, [9.159, 4.084, 2.591, 0.708], { geom: 'line', line: { color: RED, width: 2 }, flipV: 1 });
}

// ------------------------------------------------------------------------
function slide17(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [1.25, 2.558, 3.455, 3.455],
    { pts: BLOB, fill: { color: 'FFBEC0', transparency: 80 }, line: { color: RED, width: 1.5, transparency: 60, dashType: 'sysDash' } });
  sh(s, [3.038, 4.916, 4.117, 0.657],
    { geom: 'roundRect', radius: 0.5, fill: RED_B, shadow: { blur: 40, offset: 18, angle: 45, color: BLACK, opacity: 0.23 } });
  txt(s, [5.684, 5.076, 1.287, 0.37], { sz: 16, bold: 1, head: 1, align: 'right', nowrap: 1 }, ['Best Practices']);
  sh(s, [3.038, 4.286, 4.117, 0.657],
    { geom: 'roundRect', radius: 0.5, fill: CORAL, shadow: { blur: 40, offset: 18, angle: 45, color: BLACK, opacity: 0.23 } });
  txt(s, [6.234, 4.446, 0.702, 0.37], { sz: 16, bold: 1, head: 1, align: 'right', nowrap: 1 }, ['Impact']);
  sh(s, [2.19, 3.628, 4.964, 0.657],
    { geom: 'roundRect', radius: 0.5, fill: PINK, shadow: { blur: 40, offset: 18, angle: 45, color: BLACK, opacity: 0.23 } });
  txt(s, [5.649, 3.789, 1.259, 0.37], { sz: 16, bold: 1, head: 1, align: 'right', nowrap: 1 }, ['Key Principles']);
  sh(s, [3.038, 2.971, 4.117, 0.657],
    { geom: 'roundRect', radius: 0.5, fill: RED_L, shadow: { blur: 40, offset: 18, angle: 45, color: BLACK, opacity: 0.23 } });
  txt(s, [5.858, 3.131, 0.931, 0.37], { sz: 16, bold: 1, head: 1, align: 'right', nowrap: 1 }, ['Definition']);
  sh(s, [1.714, 2.986, 2.587, 2.587],
    { pts: BLOB, fill: RED, shadow: { blur: 40, offset: 18, angle: 45, color: BLACK, opacity: 0.23 } });
  txt(s, [3.208, 0.439, 6.917, 0.417], { sz: 14, align: 'center', lh: 1.5 }, ['Subtitle goes here']);
  txt(s, [1.667, 0.91, 10, 0.774], { sz: 40, bold: 1, head: 1, align: 'center' }, ['Wifi router for tourist']);
  sh(s, [7.844, 3.131, 0.72, 0.72],
    { pts: BLOB, fill: RED, shadow: { blur: 45, offset: 30, angle: 90, color: CHARCOAL, opacity: 0.28 } });
  txt(s, [7.812, 4.067, 4.452, 1.114], { sz: 14, lh: 1.5 },
    ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ']);
  txt(s, [8.849, 3.215, 3.931, 0.518], { bold: 1, head: 1, lh: 1.5 }, ['All-in-one platform']);
  sh(s, [8.038, 3.309, 0.331, 0.349], { pts: WIFI, fill: WHITE });
  sh(s, [2.519, 3.672, 0.977, 1.03], { pts: WIFI, fill: WHITE });
}

// ------------------------------------------------------------------------
function slide18(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [4.264, 1.618, 8.927, 5.716], { geom: 'roundRect', radius: 0.021, fill: CHARCOAL, flipH: 1 });
  sh(s, [0.17, 1.618, 3.941, 5.716], { geom: 'roundRect', radius: 0.021, fill: CHARCOAL, flipH: 1 });
  txt(s, [6.462, 2.611, 3.403, 2.977],
    { pts: [
    ['M', 0.001, 0.454], ['L', 1, 0], ['L', 1, 1], ['L', 0, 0.997],
    ['C', 0, 0.816, 0.001, 0.635, 0.001, 0.454], ['Z']
  ], fill: { color: 'DC8082', transparency: 50 }, shadow: { blur: 3, offset: 1, angle: 90, color: BLACK, opacity: 0.15 }, sz: 28, head: 1, align: 'center', lh: 1.3, margin: [14.4, 14.4, 0, 28.8] },
    [{ t: ' ', before: 10 }]);
  sh(s, [5.646, 5.805, 5.804, 0.783],
    { geom: 'notchedRightArrow', adj: 0.5, fill: { color: 'FFBEC0', transparency: 80 }, line: { color: RED, width: 1.5, transparency: 60, dashType: 'sysDash' } });
  txt(s, [5.398, 3.611, 1.977, 1.977],
    { pts: BLOB, fill: BLACK, shadow: { blur: 45, offset: 31, angle: 45, color: WHITE, opacity: 0.14 }, sz: 24, head: 1, align: 'center', vc: 1 },
    ['2025']);
  txt(s, [8.596, 2.611, 2.977, 2.977],
    { pts: BLOB, fill: RED, shadow: { blur: 40, offset: 18, angle: 45, color: BLACK, opacity: 0.23 }, sz: 36, bold: 1, head: 1, align: 'center', vc: 1 },
    ['34.2', 'Million']);
  sh(s, [1.418, 2.576, 1.225, 1.035],
    { pts: CAPSULE, fill: BLACK, shadow: { blur: 95, offset: 69, angle: 135, color: BLACK, opacity: 0.1 } });
  txt(s, [1.284, 2.072, 1.494, 0.37], { sz: 16, bold: 1, head: 1, align: 'center' }, ['Option']);
  txt(s, [1.418, 2.731, 1.225, 0.707], { sz: 36, bold: 1, head: 1, col: RED, align: 'center' }, ['17']);
  sh(s, [1.418, 4.245, 1.225, 1.035],
    { pts: CAPSULE, fill: BLACK, shadow: { blur: 95, offset: 69, angle: 135, color: BLACK, opacity: 0.1 } });
  txt(s, [1.284, 3.741, 1.494, 0.37], { sz: 16, bold: 1, head: 1, align: 'center' }, ['Option']);
  txt(s, [1.418, 4.4, 1.225, 0.707], { sz: 36, bold: 1, head: 1, col: RED_L, align: 'center' }, ['08']);
  sh(s, [1.418, 5.914, 1.225, 1.035],
    { pts: CAPSULE, fill: BLACK, shadow: { blur: 95, offset: 69, angle: 135, color: BLACK, opacity: 0.1 } });
  txt(s, [1.284, 5.41, 1.494, 0.37], { sz: 16, bold: 1, head: 1, align: 'center' }, ['Option']);
  txt(s, [1.418, 6.069, 1.225, 0.707], { sz: 36, bold: 1, head: 1, col: PINK, align: 'center' }, ['45']);
  txt(s, [3.208, 0.163, 6.917, 0.417], { sz: 14, align: 'center', lh: 1.5 }, ['Subtitle goes here']);
  txt(s, [1.667, 0.634, 10, 0.774], { sz: 40, bold: 1, head: 1, align: 'center' }, ['Tourist rate 2025']);
  txt(s, [7.055, 5.915, 2.217, 0.518], { bold: 1, head: 1, lh: 1.5 }, ['2025 August – End 2029']);
}

// ------------------------------------------------------------------------
function slide19(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  txt(s, [1.207, 5.359, 3.041, 0.675], { sz: 12, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR2]);
  txt(s, [1.207, 4.929, 2.462, 0.404], { bold: 1, head: 1 }, ['Definition']);
  sh(s, [9.333, 5.326, 0.792, 0],
    { geom: 'line', line: { color: WHITE, width: 1, transparency: 50, endArrowType: 'oval' } });
  txt(s, [3.208, 0.392, 6.917, 0.417], { sz: 14, align: 'center', lh: 1.5 }, ['Subtitle goes here']);
  txt(s, [1.667, 0.863, 10, 0.774], { sz: 40, bold: 1, head: 1, align: 'center' }, ['Important points for tourists']);
  txt(s, [5.331, 2.958, 2.672, 2.672],
    { pts: BLOB, fill: RED, shadow: { blur: 45, offset: 30, angle: 90, color: WHITE, opacity: 0.3 }, sz: 54, head: 1, align: 'center', vc: 1 },
    ['80%']);
  txt(s, [3.872, 4.389, 1.874, 1.874],
    { pts: BLOB, fill: { color: RED_L, transparency: 26 }, shadow: { blur: 45, offset: 30, angle: 90, color: WHITE, opacity: 0.3 }, sz: 24, head: 1, align: 'center', vc: 1 },
    ['20%']);
  txt(s, [7.433, 4.611, 1.652, 1.652],
    { pts: BLOB, fill: { color: PINK, transparency: 26 }, shadow: { blur: 45, offset: 30, angle: 90, color: WHITE, opacity: 0.3 }, sz: 28, head: 1, align: 'center', vc: 1 },
    ['30%']);
  txt(s, [5.156, 2.282, 1.188, 1.188],
    { pts: BLOB, fill: { color: CORAL, transparency: 26 }, shadow: { blur: 45, offset: 30, angle: 90, color: WHITE, opacity: 0.3 }, sz: 28, head: 1, align: 'center', vc: 1 },
    ['40%']);
  sh(s, [8.241, 3.416, 0.792, 0],
    { geom: 'line', line: { color: WHITE, width: 1, transparency: 50, endArrowType: 'oval' } });
  sh(s, [4.017, 2.654, 0.792, 0],
    { geom: 'line', line: { color: WHITE, width: 1, transparency: 50, endArrowType: 'oval' }, flipH: 1 });
  sh(s, [2.949, 5.123, 0.792, 0],
    { geom: 'line', line: { color: WHITE, width: 1, transparency: 50, endArrowType: 'oval' }, flipH: 1 });
  txt(s, [10.3, 5.504, 3.041, 0.675], { sz: 12, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR2]);
  txt(s, [10.3, 5.074, 2.462, 0.404], { bold: 1, head: 1 }, ['Impact']);
  txt(s, [9.271, 3.64, 3.041, 0.675], { sz: 12, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR2]);
  txt(s, [9.271, 3.21, 2.462, 0.404], { bold: 1, head: 1 }, ['Key Principles']);
  txt(s, [1.561, 2.893, 3.041, 0.675], { sz: 12, lh: 1.5 }, [T_LOREM_IPSUM_DOLOR2]);
  txt(s, [1.561, 2.463, 2.462, 0.404], { bold: 1, head: 1 }, ['Best Practices']);
}

// ------------------------------------------------------------------------
function slide20(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [0, 1.696, 9.41, 5.804],
    { pts: [
    ['M', 0.56, 0], ['C', 0.19, 0, 0.036, 0.175, 0.004, 0.677], ['L', 0, 0.757], ['L', 0.004, 0.837],
    ['C', 0.007, 0.884, 0.011, 0.928, 0.016, 0.97], ['L', 0.02, 1], ['L', 1, 1], ['L', 1, 0.201],
    ['L', 0.995, 0.193], ['C', 0.907, 0.055, 0.767, 0, 0.56, 0], ['Z']
  ], fill: PINK, rot: 180, flipV: 1 });
  sh(s, [0, 3.351, 8.088, 4.149],
    { pts: [
    ['M', 0.569, 0], ['C', 0.194, 0, 0.037, 0.214, 0.004, 0.828], ['L', 0, 0.926], ['L', 0.003, 1],
    ['L', 1, 1], ['L', 1, 0.216], ['L', 0.98, 0.184], ['C', 0.89, 0.053, 0.757, 0, 0.569, 0], ['Z']
  ], fill: RED_L, rot: 180, flipV: 1 });
  sh(s, [0, 4.73, 7.03, 2.77],
    { pts: [
    ['M', 0.575, 0], ['C', 0.213, 0, 0.05, 0.251, 0.002, 0.959], ['L', 0, 1], ['L', 1, 1], ['L', 1, 0.252],
    ['L', 0.997, 0.245], ['C', 0.904, 0.071, 0.768, 0, 0.575, 0], ['Z']
  ], fill: RED, rot: 180, flipV: 1 });
  txt(s, [2.156, 5.56, 2.891, 1.111], { sz: 16, lh: 1.3 }, [T_LOREM_IPSUM_DECIDED]);
  txt(s, [7.03, 1.07, 5.423, 1.33], { sz: 40, bold: 1, head: 1, align: 'right', lh: 0.9 },
    ['Create experience', 'With us']);
  txt(s, [9.603, 0.602, 2.85, 0.417], { sz: 14, align: 'right', lh: 1.5 }, ['Subtitle goes here']);
  txt(s, [2.271, 3.75, 3.547, 0.767], { sz: 16, lh: 1.3 }, [T_LOREM_IPSUM_DECIDED]);
  txt(s, [3.208, 2.236, 3.547, 0.767], { sz: 16, lh: 1.3 }, [T_LOREM_IPSUM_DECIDED]);
  txt(s, [1.846, 2.081, 1.362, 0.909], { sz: 48, bold: 1, head: 1, align: 'center' }, ['01']);
  txt(s, [0.734, 3.695, 1.362, 0.909], { sz: 48, bold: 1, head: 1, align: 'center' }, ['02']);
  txt(s, [0.719, 5.529, 1.362, 0.909], { sz: 48, bold: 1, head: 1, align: 'center' }, ['03']);
}

// ------------------------------------------------------------------------
function slide21(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [2.292, 0, 11.042, 7.5],
    { geom: 'homePlate', adj: 0.167, fill: CHARCOAL, shadow: { blur: 100, offset: 200, angle: 180, color: BLACK, opacity: 0.4 }, flipH: 1 });
  sh(s, [9.609, 0, 3.724, 7.5],
    { geom: 'homePlate', adj: 1, fill: RED, shadow: { blur: 100, offset: 75, angle: 180, color: BLACK, opacity: 0.4 }, flipH: 1 });
  photo(s, [0, 0, 5.991, 7.5]);
  txt(s, [10.773, 3.144, 1.859, 1.212], { sz: 66, head: 1, align: 'right' }, ['02']);
  txt(s, [5.164, 2.437, 4.463, 1.111], { sz: 60, head: 1 }, ['Section Break']);
  txt(s, [5.207, 2.101, 3.478, 0.337], { sz: 14, head: 1 }, ['Subtitle Here']);
  txt(s, [5.207, 4.692, 1.672, 0.462],
    { geom: 'roundRect', radius: 0.5, fill: RED, shadow: { blur: 30, offset: 15, angle: 90, color: BLACK, opacity: 0.1 }, sz: 14, head: 1, align: 'center', vc: 1 },
    ['Subtitle Here']);
}

// ------------------------------------------------------------------------
function slide22(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [6.745, 1.294, 2.087, 4.897], { geom: 'roundRect', radius: 0.1, fill: CHARCOAL });
  sh(s, [8.997, 1.294, 2.087, 4.897], { geom: 'roundRect', radius: 0.1, fill: CHARCOAL });
  sh(s, [1.115, 5.128, 5.474, 1.063], { geom: 'roundRect', radius: 0.1, fill: RED });
  txt(s, [1.273, 5.588, 2.087, 0.274], { sz: 7, ital: 1, lh: 1.5 }, ['Photographer']);
  txt(s, [1.273, 5.342, 2.087, 0.286], { sz: 11, head: 1 }, ['Joule Jessa ']);
  txt(s, [5.938, 5.342, 0.48, 0.48],
    { geom: 'roundRect', radius: 0.1, fill: RED_L, shadow: { blur: 18, offset: 6, angle: 45, color: BLACK, opacity: 0.19 }, sz: 12, bold: 1, head: 1, align: 'center', vc: 1 },
    ['01']);
  txt(s, [6.909, 4.246, 1.733, 0.274], { sz: 7, ital: 1, lh: 1.5 }, ['Photographer']);
  txt(s, [6.909, 4, 1.733, 0.286], { sz: 11, head: 1 }, ['Joule Jessa ']);
  txt(s, [8.161, 5.471, 0.48, 0.48],
    { geom: 'roundRect', radius: 0.1, fill: RED, shadow: { blur: 18, offset: 6, angle: 45, color: BLACK, opacity: 0.19 }, sz: 12, bold: 1, head: 1, align: 'center', vc: 1 },
    ['02']);
  txt(s, [10.445, 5.471, 0.48, 0.48],
    { geom: 'roundRect', radius: 0.1, fill: RED, shadow: { blur: 18, offset: 6, angle: 45, color: BLACK, opacity: 0.19 }, sz: 12, bold: 1, head: 1, align: 'center', vc: 1 },
    ['03']);
  txt(s, [9.193, 4.246, 1.733, 0.274], { sz: 7, ital: 1, lh: 1.5 }, ['Photographer']);
  txt(s, [9.193, 4, 1.733, 0.286], { sz: 11, head: 1 }, ['Joule Jessa ']);
  photo(s, [1.115, 1.269, 5.474, 3.662],
    { pts: [
    ['M', 0.033, 0], ['L', 0.967, 0], ['C', 0.985, 0, 1, 0.022, 1, 0.05], ['L', 1, 0.95],
    ['C', 1, 0.978, 0.985, 1, 0.967, 1], ['L', 0.033, 1], ['C', 0.015, 1, 0, 0.978, 0, 0.95],
    ['L', 0, 0.05], ['C', 0, 0.022, 0.015, 0, 0.033, 0], ['Z']
  ] });
  photo(s, [6.75, 1.294, 2.082, 2.456],
    { pts: [
    ['M', 0.05, 0], ['L', 0.95, 0], ['C', 0.978, 0, 1, 0.019, 1, 0.042], ['L', 1, 0.958],
    ['C', 1, 0.981, 0.978, 1, 0.95, 1], ['L', 0.05, 1], ['C', 0.022, 1, 0, 0.981, 0, 0.958],
    ['L', 0, 0.042], ['C', 0, 0.019, 0.022, 0, 0.05, 0], ['Z']
  ] });
  photo(s, [8.99, 1.294, 2.087, 2.456],
    { pts: [
    ['M', 0.05, 0], ['L', 0.95, 0], ['C', 0.978, 0, 1, 0.019, 1, 0.042], ['L', 1, 0.958],
    ['C', 1, 0.981, 0.978, 1, 0.95, 1], ['L', 0.05, 1], ['C', 0.022, 1, 0, 0.981, 0, 0.958],
    ['L', 0, 0.042], ['C', 0, 0.019, 0.022, 0, 0.05, 0], ['Z']
  ] });
}

// ------------------------------------------------------------------------
function slide23(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  txt(s, [1.115, 1.073, 8.665, 1.313], { sz: 72, head: 1 }, ['Awesome Team']);
  txt(s, [11.957, 4.746, 1.303, 0.252], { rot: 90, sz: 9, align: 'center' }, ['PRESENTATION']);
  txt(s, [11.964, 1.331, 1.303, 0.252], { rot: 90, sz: 9, align: 'center' }, ['TEAM']);
  txt(s, [1.273, 3.295, 2.087, 0.274], { sz: 7, ital: 1, lh: 1.5 }, ['Photographer']);
  txt(s, [1.273, 3.049, 2.087, 0.286], { sz: 11, head: 1 }, ['Joule Jessa ']);
  txt(s, [4.786, 3.295, 1.733, 0.274], { sz: 7, ital: 1, lh: 1.5 }, ['Photographer']);
  txt(s, [4.786, 3.049, 1.733, 0.286], { sz: 11, head: 1 }, ['Joule Jessa ']);
  txt(s, [9.193, 3.295, 1.733, 0.274], { sz: 7, ital: 1, lh: 1.5 }, ['Photographer']);
  txt(s, [9.193, 3.049, 1.733, 0.286], { sz: 11, head: 1 }, ['Joule Jessa ']);
  photo(s, [1.115, 3.75, 3.189, 3.425]);
  photo(s, [4.462, 3.75, 4.37, 3.425]);
  photo(s, [8.99, 3.75, 3.229, 3.425]);
}

// ------------------------------------------------------------------------
function slide24(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  photo(s, [6.745, 0, 6.589, 6.154],
    { pts: [
    ['M', 0.047, 0], ['L', 0.953, 0], ['C', 0.979, 0, 1, 0.022, 1, 0.05], ['L', 1, 0.95],
    ['C', 1, 0.978, 0.979, 1, 0.953, 1], ['L', 0.047, 1], ['C', 0.021, 1, 0, 0.978, 0, 0.95],
    ['L', 0, 0.05], ['C', 0, 0.022, 0.021, 0, 0.047, 0], ['Z']
  ] });
  txt(s, [0.776, 1.137, 5.553, 0.841], { sz: 44, head: 1 }, ['Amanda kelsey']);
  txt(s, [0.776, 2.004, 4.168, 0.421], { sz: 19, head: 1 }, ['Photographer']);
  txt(s, [0.776, 3.133, 4.131, 0.802], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS2, T_DERITI_VOLS_ESSE]);
  txt(s, [0.776, 2.846, 3.478, 0.286], { sz: 11, head: 1 }, ['Modern Design']);
  sh(s, [0.87, 5.118, 4.759, 0.303],
    { fill: BLACK, shadow: { blur: 18, offset: 6, angle: 45, color: BLACK, opacity: 0.12 } });
  txt(s, [0.87, 5.118, 2.605, 0.303],
    { fill: { color: WHITE, transparency: 80 }, sz: 9, bold: 1, head: 1, align: 'right', vc: 1, margin: [11.3, 11.3, 7.2, 7.2] },
    ['45%']);
  txt(s, [1.068, 5.14, 0.857, 0.265], { sz: 9, bold: 1, head: 1, lh: 1.3, nowrap: 1, margin: [0, 17, 2.8, 2.8] },
    [{ t: 'Your Text Here', before: 10 }]);
  sh(s, [0.87, 4.648, 4.759, 0.303],
    { fill: BLACK, shadow: { blur: 18, offset: 6, angle: 45, color: BLACK, opacity: 0.12 } });
  txt(s, [0.87, 4.648, 4.39, 0.303],
    { fill: { color: RED, transparency: 43 }, sz: 9, bold: 1, head: 1, align: 'right', vc: 1, margin: [11.3, 11.3, 7.2, 7.2] },
    ['90%']);
  txt(s, [1.068, 4.669, 0.857, 0.265], { sz: 9, bold: 1, head: 1, lh: 1.3, nowrap: 1, margin: [0, 17, 2.8, 2.8] },
    [{ t: 'Your Text Here', before: 10 }]);
  sh(s, [0.87, 4.178, 4.759, 0.303],
    { fill: BLACK, shadow: { blur: 18, offset: 6, angle: 45, color: BLACK, opacity: 0.12 } });
  txt(s, [0.87, 4.178, 2.605, 0.303],
    { fill: { color: WHITE, transparency: 80 }, sz: 9, bold: 1, head: 1, align: 'right', vc: 1, margin: [11.3, 11.3, 7.2, 7.2] },
    ['45%']);
  txt(s, [1.068, 4.199, 0.857, 0.265], { sz: 9, bold: 1, head: 1, lh: 1.3, nowrap: 1, margin: [0, 17, 2.8, 2.8] },
    [{ t: 'Your Text Here', before: 10 }]);
  sh(s, [12.676, 0.623, 0.329, 0.329], { pts: ARROW_NE, fill: WHITE });
  sh(s, [6.195, 5.572, 1.099, 0.791], { geom: 'roundRect', radius: 0.1, fill: RED });
  sh(s, [6.667, 5.806, 0.323, 0.323], { pts: ARROW_NE, fill: WHITE });
}

// ------------------------------------------------------------------------
function slide25(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  photo(s, [5.445, 0.592, 3.172, 6.315], { device: 'phone' });
  txt(s, [1.027, 4.152, 3.435, 1.043], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS3]);
  txt(s, [0.984, 1.256, 4.463, 1.767], { sz: 33, head: 1 }, ['Application', 'Devices', 'Mobile']);
  txt(s, [1.027, 3.866, 3.478, 0.286], { sz: 11, head: 1 }, ['Modern Design']);
  txt(s, [0.984, 5.854, 1.651, 0.356], { sz: 14, und: 1 }, ['Button Here']);
  txt(s, [9.514, 5.854, 1.651, 0.356], { sz: 14, und: 1 }, ['Button Here']);
  txt(s, [9.514, 3.61, 1.593, 0.286], { sz: 11 }, ['TEXT TITTLE HERE']);
  txt(s, [9.514, 3.896, 2.644, 0.986], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS]);
  txt(s, [9.514, 1.751, 1.593, 0.286], { sz: 11 }, ['TEXT TITTLE HERE']);
  txt(s, [9.514, 2.037, 2.644, 0.986], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS]);
  photo(s, [5.748, 0.971, 2.586, 5.555],
    { pts: [
    ['M', 0.131, 0], ['L', 0.869, 0], ['C', 0.942, 0, 1, 0.027, 1, 0.061], ['L', 1, 0.939],
    ['C', 1, 0.973, 0.942, 1, 0.869, 1], ['L', 0.131, 1], ['C', 0.058, 1, 0, 0.973, 0, 0.939],
    ['L', 0, 0.061], ['C', 0, 0.027, 0.058, 0, 0.131, 0], ['Z']
  ] });
}

// ------------------------------------------------------------------------
function slide26(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [5.304, -0.008, 3.736, 7.508], { geom: 'roundRect', radius: 0, fill: BLACK });
  sh(s, [1.115, 4.812, 3.387, 0.303], { fill: 'F6F5EF' });
  txt(s, [1.115, 4.812, 1.854, 0.303],
    { fill: RED_L, sz: 9, bold: 1, head: 1, align: 'right', vc: 1, margin: [11.3, 11.3, 7.2, 7.2] }, ['45%']);
  txt(s, [1.255, 4.833, 0.857, 0.265], { sz: 9, bold: 1, head: 1, lh: 1.3, nowrap: 1, margin: [0, 17, 2.8, 2.8] },
    [{ t: 'Your Text Here', before: 10 }]);
  sh(s, [1.115, 5.747, 3.387, 0.303], { fill: 'F6F5EF' });
  txt(s, [1.115, 5.747, 3.124, 0.303],
    { fill: BLACK, sz: 9, bold: 1, head: 1, col: BLACK, align: 'right', vc: 1, margin: [11.3, 11.3, 7.2, 7.2] },
    ['90%']);
  txt(s, [1.255, 5.768, 0.857, 0.265],
    { sz: 9, bold: 1, head: 1, col: BLACK, lh: 1.3, nowrap: 1, margin: [0, 17, 2.8, 2.8] },
    [{ t: 'Your Text Here', before: 10 }]);
  sh(s, [1.115, 5.28, 3.387, 0.303],
    { fill: BLACK, shadow: { blur: 18, offset: 6, angle: 45, color: BLACK, opacity: 0.12 } });
  txt(s, [1.115, 5.28, 2.38, 0.303],
    { fill: RED, sz: 9, bold: 1, head: 1, align: 'right', vc: 1, margin: [11.3, 11.3, 7.2, 7.2] }, ['60%']);
  txt(s, [1.255, 5.301, 0.857, 0.265], { sz: 9, bold: 1, head: 1, lh: 1.3, nowrap: 1, margin: [0, 17, 2.8, 2.8] },
    [{ t: 'Your Text Here', before: 10 }]);
  txt(s, [1.115, 3.733, 3.657, 0.8], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS]);
  txt(s, [1.115, 1.256, 4.463, 1.767], { sz: 33, head: 1 }, ['Application', 'Devices', 'Smartwatch']);
  photo(s, [10.085, 1.617, 2.703, 4.232], { device: 'watchAlpine' });
  photo(s, [5.833, 1.617, 2.657, 4.232], { device: 'watchSport' });
  sh(s, [8.832, 3.224, 0.984, 0.984],
    { geom: 'ellipse', fill: RED, shadow: { blur: 18, offset: 6, angle: 45, color: BLACK, opacity: 0.12 } });
  txt(s, [8.922, 3.429, 0.804, 0.438], { fill: RED, sz: 20, bold: 1, head: 1, align: 'center' }, ['VS']);
  txt(s, [8.967, 3.742, 0.714, 0.271], { fill: RED, sz: 9, align: 'center', lh: 1.2 }, ['Point']);
  photo(s, [6.188, 2.562, 1.951, 2.41],
    { pts: [
    ['M', 0.215, 0], ['L', 0.785, 0], ['C', 0.904, 0, 1, 0.078, 1, 0.174], ['L', 1, 0.826],
    ['C', 1, 0.922, 0.904, 1, 0.785, 1], ['L', 0.215, 1], ['C', 0.096, 1, 0, 0.922, 0, 0.826],
    ['L', 0, 0.174], ['C', 0, 0.078, 0.096, 0, 0.215, 0], ['Z']
  ] });
  photo(s, [10.521, 2.604, 1.855, 2.278],
    { pts: [
    ['M', 0.265, 0], ['L', 0.735, 0], ['C', 0.881, 0, 1, 0.097, 1, 0.216], ['L', 1, 0.784],
    ['C', 1, 0.903, 0.881, 1, 0.735, 1], ['L', 0.265, 1], ['C', 0.118, 1, 0, 0.903, 0, 0.784],
    ['L', 0, 0.216], ['C', 0, 0.097, 0.118, 0, 0.265, 0], ['Z']
  ] });
}

// ------------------------------------------------------------------------
function slide27(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  txt(s, [1.027, 5.415, 3.435, 0.322], { sz: 9, lh: 1.5 }, ['Lorep  ipsum duis aute irure dolor in kauselih oilue']);
  txt(s, [0.984, 1.256, 4.463, 1.767], { sz: 33, head: 1 }, ['Application', 'Devices', 'Tablet']);
  txt(s, [1.027, 5.128, 3.478, 0.286], { sz: 11, head: 1 }, ['Modern Design']);
  photo(s, [5.608, 0.15, 6.613, 8.461], { device: 'tablet' });
  photo(s, [6.104, 0.646, 5.603, 7.444],
    { pts: [
    ['M', 0.022, 0], ['L', 0.978, 0], ['C', 0.99, 0, 1, 0.008, 1, 0.017], ['L', 1, 0.983],
    ['C', 1, 0.992, 0.99, 1, 0.978, 1], ['L', 0.022, 1], ['C', 0.01, 1, 0, 0.992, 0, 0.983],
    ['L', 0, 0.017], ['C', 0, 0.008, 0.01, 0, 0.022, 0], ['Z']
  ] });
}

// ------------------------------------------------------------------------
function slide28(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  txt(s, [6.667, 1.256, 4.463, 1.212], { sz: 33, head: 1 }, ['Application', 'Devices Laptop']);
  txt(s, [6.71, 5.854, 1.651, 0.356], { sz: 14, und: 1 }, ['Button Here']);
  photo(s, [-4.814, -0.038, 11.856, 7.334], { device: 'laptop' });
  txt(s, [9.514, 3.61, 1.593, 0.286], { sz: 11 }, ['TEXT TITTLE HERE']);
  txt(s, [9.514, 3.896, 2.644, 0.986], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS]);
  txt(s, [6.71, 3.61, 1.593, 0.286], { sz: 11 }, ['TEXT TITTLE HERE']);
  txt(s, [6.71, 3.896, 2.644, 0.986], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS]);
  photo(s, [-3.281, 0.875, 8.8, 5.514]);
}

// ------------------------------------------------------------------------
function slide29(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  photo(s, [2.217, 1.269, 10.946, 6.064]);
  sh(s, [0.957, 0.325, 4.49, 4.601], { geom: 'roundRect', radius: 0.1, fill: CHARCOAL });
  txt(s, [10.132, 4.335, 1.593, 0.286], { sz: 11 }, ['BOOK A CALL']);
  txt(s, [10.132, 4.621, 1.593, 0.305], { sz: 9, lh: 1.5 }, ['(+12) 345 678 90123']);
  txt(s, [10.132, 5.569, 2.458, 0.286], { sz: 11 }, ['EMAIL US']);
  txt(s, [10.132, 5.855, 2.458, 0.305], { sz: 9, lh: 1.5 }, ['yourawesome@mail.com']);
  txt(s, [10.132, 2.874, 1.593, 0.286], { sz: 11 }, ['MEET US']);
  txt(s, [10.132, 3.16, 1.593, 0.532], { sz: 9, lh: 1.5 }, ['123, Sesame street, Malaysia']);
  txt(s, [10.132, 1.414, 1.593, 0.286], { sz: 11 }, ['FOLLOW US']);
  txt(s, [10.132, 1.7, 1.593, 0.532], { sz: 9, lh: 1.5 }, ['@yourbrandhere', '@anotherbrand']);
  txt(s, [1.53, 1.504, 3.343, 2.121], { sz: 60, head: 1, align: 'center' }, ['Get In Touch']);
  sh(s, [0.957, 5.128, 4.49, 2.029], { geom: 'roundRect', radius: 0.1, fill: RED });
  txt(s, [1.406, 5.62, 2.164, 0.286], { sz: 11 }, ['TEXT TITTLE HERE']);
  txt(s, [1.406, 5.906, 3.591, 0.759], { sz: 9, lh: 1.5 }, [T_LOREP_IPSUM_DUIS]);
}

// ------------------------------------------------------------------------
function slide30(pptx) {
  const s = pptx.addSlide();
  s.background = { color: BLACK };
  sh(s, [0.957, 4.066, 1.26, 2.008], { geom: 'roundRect', radius: 0.5, fill: RED });
  sh(s, [2.217, 4.066, 2.245, 2.008], { geom: 'roundRect', radius: 0.5, fill: RED_L });
  sh(s, [4.422, 4.066, 6.812, 2.008], { geom: 'roundRect', radius: 0.5, fill: CHARCOAL });
  sh(s, [11.234, 4.066, 2.099, 2.008], { geom: 'roundRect', radius: 0.5, fill: CORAL });
  txt(s, [0.957, 1.073, 8.665, 1.582], { sz: 88, head: 1 }, ['Thank You.']);
  txt(s, [1.22, 4.741, 0.879, 0.656], { sz: 32, head: 1 }, ['30']);
  txt(s, [11.019, 2.065, 0.984, 0.286], { sz: 11 }, ['End Slide']);
  txt(s, [9.014, 2.081, 1.303, 0.252], { sz: 9 }, ['MODERN DESIGN']);
  txt(s, [5.447, 4.912, 4.53, 0.723], { sz: 8, lh: 1.5 },
    [{ r: ['Lorep  ipsum duis ', { t: 'aute irure ', bold: 1 }, T_DOLOR_IN_KAUSELIH] }, { r: [T_ESSE_CILL_INURE, { t: ' irusitakus ', bold: 1 }, 'reprei'] }, 'Voluptate lorem kuisais.']);
  txt(s, [0.957, 2.859, 2.947, 0.337], { sz: 14, head: 1 }, ['Presentation Design']);
  txt(s, [5.447, 4.489, 1.303, 0.252], { sz: 9 }, ['MODERN DESIGN']);
}

// ------------------------------------------------------------------------
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.title = 'Travel Tourism Deck';
  [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11,
    slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22,
    slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30
  ].forEach(fn => fn(pptx));
  return pptx.writeFile({ fileName: path.join(__dirname, '0010fd63-2868-4489-9e92-678a431e9e89_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
