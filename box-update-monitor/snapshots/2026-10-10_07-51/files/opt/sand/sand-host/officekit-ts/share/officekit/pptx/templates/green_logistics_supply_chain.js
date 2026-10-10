/**
 * Green Logistic — presentation template (30 slides, 13.333 x 7.5 in).
 *
 * Standalone pptxgenjs rebuild of the reference deck.  Photographs and
 * vector icons from the original are replaced with flat colour placeholders
 * (see `icon()` / `photo()`), everything else is drawn with native shapes.
 *
 *   node <thisfile>.js   ->  writes the .pptx next to the script
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ── palette ─────────────────────────────────────────────────────────────────
const GREEN      = '228B22';   // accent1
const GREEN_D    = '196819';   // accent1, 75% luminance
const WM_GREEN   = '5DD75D';   // watermark outline (accent1 light)
const ORANGE     = 'FFA500';   // accent2
const WM_ORANGE  = 'FFC966';   // watermark outline (accent2 light)
const CREAM      = 'FFEDCC';   // accent2, 20% lum / 80% off
const WHITE      = 'FFFFFF';
const INK        = '000000';
const GRAY       = '595959';   // body copy on light backgrounds
const GRAY_D     = '404040';
const GRAY_M     = '666666';

const HEAD = 'Montserrat Bold';   // theme major font
const BODY = 'Roboto Regular';    // theme minor font

// ── reusable shadow presets (from the template's effect list) ───────────────
const SH_TIGHT = { type: 'outer', color: INK, opacity: 0.10, blur: 65, offset: 12, angle: 146 };
const SH_CARD  = { type: 'outer', color: INK, opacity: 0.20, blur: 65, offset: 12, angle: 146 };
const SH_SOFT  = { type: 'outer', color: INK, opacity: 0.10, blur: 65, offset: 31, angle: 146 };
const SH_DEEP  = { type: 'outer', color: INK, opacity: 0.48, blur: 65, offset: 31, angle: 146 };

// ── boilerplate text that repeats all over the deck ─────────────────────────
const L1 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.';
const L2 = L1 + ' Aenean commodo ligula eget dolor. Aenean massa.';
const L3 = L1 + ' Aenean commodo ligula eget dolor. ';
const L4 = L1 + ' Aenean commodo ligula.';
const L5 = 'Lorem ipsum dolor sit amet, consectetuer.';
const L6 = L2 + ' Cum sociis natoque penatibus et magnis dis parturient montes. ';
const L7 = L1 + ' Aenean commodo ligula eget.';
const LQ = '\u201CLorem ipsum dolor sit amet, consectetuer adipiscing elit.\u201D';

// ── low-level drawing helpers ───────────────────────────────────────────────

/** Solid shape.  `o` = {a:alpha%, r:corner radius in, ln/lw/la:outline, sh:shadow, rot:deg} */
function box(s, geom, x, y, w, h, color, o) {
  o = o || {};
  const opt = { x: x, y: y, w: w, h: h };
  if (color) opt.fill = o.a ? { color: color, transparency: o.a } : { color: color };
  else opt.fill = { type: 'none' };
  if (o.ln) opt.line = { color: o.ln, width: o.lw || 1, transparency: o.la };
  else if (geom !== 'line') opt.line = { type: 'none' };
  if (geom === 'roundRect') opt.rectRadius = o.r;
  if (o.sh) opt.shadow = o.sh;
  if (o.rot) opt.rotate = o.rot;
  s.addShape(geom, opt);
  return s;
}

/** Text frame.  `o` = {al,va,sz,c,hd,ff,b,i,ls,sb,vert,nowrap,hollow,ln,la} */
function tx(s, text, x, y, w, h, o) {
  o = o || {};
  const opt = {
    x: x, y: y, w: w, h: h,
    fontSize: o.sz || 18,
    fontFace: o.hd ? HEAD : (o.ff || BODY),
    color: o.c || INK,
    align: o.al || 'left',
    valign: o.va || 'top',
    margin: [7.2, 7.2, 3.6, 3.6],
    wrap: !o.nowrap,
  };
  if (o.b) opt.bold = true;
  if (o.i) opt.italic = true;
  if (o.ls) opt.lineSpacingMultiple = o.ls;
  if (o.sb) opt.paraSpaceBefore = o.sb;
  if (o.vert) opt.vert = o.vert;
  if (o.hollow) {                       // outlined "Logistic" watermark lettering
    opt.color = WHITE;
    opt.transparency = 100;
    opt.outline = { color: { color: o.ln, transparency: o.la || 0 }, size: 1 };
  }
  s.addText(text, opt);
  return s;
}

/** One styled run inside a `tx()` text array; same option keys as `tx`. */
function run(text, o) {
  o = o || {};
  const opt = {};
  if (o.sz) opt.fontSize = o.sz;
  if (o.hd) opt.fontFace = HEAD;
  else if (o.ff) opt.fontFace = o.ff;
  if (o.c) opt.color = o.c;
  if (o.b) opt.bold = true;
  if (o.i) opt.italic = true;
  if (o.brk) opt.breakLine = true;
  return { text: text, options: opt };
}

/** Placeholder for a vector icon in the reference deck. */
function icon(s, x, y, w, h, color) {
  return box(s, 'ellipse', x, y, w, h, color);
}

/** Placeholder for a raster photograph in the reference deck. */
function photo(s, x, y, w, h, color, radius) {
  box(s, radius ? 'roundRect' : 'rect', x, y, w, h, color || 'D9D9D9', { r: radius });
  tx(s, '[image]', x, y + h / 2 - 0.2, w, 0.4, { al: 'center', sz: 12, c: WHITE });
  return s;
}


// ── slide furniture: header chevrons + business name, footer url + page badge ─
// Layouts that hide the master furniture supply their own copy, hence the
// per-slide colour table below.
const CHROME = {
   1: { web: WHITE, biz: WHITE, badge: null,    page: false },
   2: { web: INK,   biz: INK,   badge: GREEN_D, page: true  },
   3: { web: WHITE, biz: WHITE, badge: ORANGE,  page: true  },
   4: { web: INK,   biz: INK,   badge: GREEN_D, page: true  },
   5: { web: INK,   biz: INK,   badge: GREEN_D, page: true  },
   6: { web: WHITE, biz: WHITE, badge: GREEN_D, page: true  },
   7: { web: INK,   biz: INK,   badge: GREEN_D, page: true  },
   8: { web: INK,   biz: INK,   badge: GREEN_D, page: true  },
   9: { web: WHITE, biz: WHITE, badge: ORANGE,  page: true  },
  10: { web: INK,   biz: INK,   badge: GREEN_D, page: true  },
  11: { web: INK,   biz: WHITE, badge: GREEN,   page: true  },
  12: { web: WHITE, biz: INK,   badge: ORANGE,  page: true  },
  13: { web: WHITE, biz: INK,   badge: GREEN,   page: true  },
  14: { web: WHITE, biz: WHITE, badge: GREEN_D, page: true  },
  15: { web: WHITE, biz: WHITE, badge: GREEN_D, page: true  },
  16: { web: WHITE, biz: INK,   badge: ORANGE,  page: true  },
  17: { web: INK,   biz: INK,   badge: GREEN_D, page: true  },
  18: { web: WHITE, biz: WHITE, badge: ORANGE,  page: true  },
  19: { web: WHITE, biz: WHITE, badge: ORANGE,  page: true  },
  20: { web: WHITE, biz: INK,   badge: ORANGE,  page: true  },
  21: { web: WHITE, biz: INK,   badge: ORANGE,  page: true  },
  22: { web: INK,   biz: INK,   badge: ORANGE,  page: true  },
  23: { web: WHITE, biz: INK,   badge: ORANGE,  page: true  },
  24: { web: INK,   biz: INK,   badge: ORANGE,  page: true  },
  25: { web: INK,   biz: INK,   badge: GREEN_D, page: true  },
  26: { web: INK,   biz: INK,   badge: GREEN,   page: true  },
  27: { web: INK,   biz: INK,   badge: ORANGE,  page: true, headY: 0.333 },
  28: { web: INK,   biz: INK,   badge: GREEN_D, page: true  },
  29: { web: WHITE, biz: WHITE, badge: GREEN_D, page: true  },
  30: { web: WHITE, biz: WHITE, badge: null,    page: false },
};

function chrome(s, n) {
  const c = CHROME[n];
  const hy = c.headY || 0.229;
  box(s, 'chevron', 0.568, hy, 0.316, 0.316, ORANGE);
  box(s, 'chevron', 0.794, hy, 0.316, 0.316, CREAM);
  tx(s, 'Business Name', 1.086, hy - 0.027, 2.966, 0.337, { sz: 14, c: c.biz });
  tx(s, 'www.yourwebsite.com', 0.559, 6.88, 2.966, 0.337, { sz: 14, c: c.web });
  if (c.badge) box(s, 'roundRect', 12.452, 6.885, 0.326, 0.326, c.badge, { r: 0.054 });
  if (c.page) tx(s, String(n), 12.289, 6.88, 0.652, 0.337, { al: 'center', sz: 14, c: WHITE });
  return s;
}

function newSlide(pres, bgColor) {
  const s = pres.addSlide();
  s.background = { color: bgColor || WHITE };
  return s;
}

/** Open circular arc (used for the donut gauges on slide 17). */
function arc(s, x, y, w, h, startDeg, sweepDeg, o) {
  s.addShape('arc', {
    x: x, y: y, w: w, h: h,
    fill: { type: 'none' },
    line: { color: o.ln, width: o.lw || 1, transparency: o.la },
    angleRange: [startDeg, (startDeg + sweepDeg) % 360],
  });
  return s;
}


// ── stylised vector illustrations ───────────────────────────────────────────
// The reference deck draws these with several hundred freeform paths; they are
// rebuilt here as a handful of polygons in true isometric projection.

const ISO_X = 0.866, ISO_Y = 0.5;   // 30-degree isometric unit vectors

/** Filled polygon from absolute inch coordinates. */
function poly(s, color, pts) {
  const xs = pts.map(p => p.x), ys = pts.map(p => p.y);
  const x0 = Math.min.apply(null, xs), y0 = Math.min.apply(null, ys);
  s.addShape('custGeom', {
    x: x0, y: y0, w: Math.max.apply(null, xs) - x0, h: Math.max.apply(null, ys) - y0,
    fill: { color: color }, line: { type: 'none' },
    points: pts.map(p => ({ x: p.x - x0, y: p.y - y0 })).concat([{ close: true }]),
  });
  return s;
}

/**
 * Isometric cuboid.  (ox,oy) is the screen position of the top-rear corner,
 * `du`/`dv` the ground footprint and `h` the extrusion height (inches).
 */
function isoBox(s, ox, oy, du, dv, h, top, left, right) {
  const P = (u, v, dz) => ({ x: ox + (u - v) * ISO_X, y: oy + (u + v) * ISO_Y + dz });
  poly(s, top,   [P(0, 0, 0), P(du, 0, 0), P(du, dv, 0), P(0, dv, 0)]);
  poly(s, left,  [P(0, dv, 0), P(du, dv, 0), P(du, dv, h), P(0, dv, h)]);
  poly(s, right, [P(du, dv, 0), P(du, 0, 0), P(du, 0, h), P(du, dv, h)]);
  return s;
}

/** Slide 24 — isometric warehouse with loading bay, truck, forklift, pallet. */
function warehouse(s) {
  // rear (taller) hall -------------------------------------------------------
  isoBox(s, 2.71, 1.99, 1.75, 1.75, 1.62, '5DD75D', '228B22', '228B22');   // green trim slab
  isoBox(s, 2.71, 1.89, 1.66, 1.66, 1.52, 'EEEEEE', 'E6E6E6', 'D1D1D1');   // walls
  poly(s, '103940', [{ x: 2.07, y: 3.35 }, { x: 2.72, y: 3.72 }, { x: 2.72, y: 5.05 }, { x: 2.07, y: 4.68 }]);
  poly(s, '228B22', [{ x: 2.20, y: 3.50 }, { x: 2.63, y: 3.75 }, { x: 2.63, y: 4.90 }, { x: 2.20, y: 4.65 }]);
  isoBox(s, 2.05, 1.62, 0.16, 0.16, 0.18, 'E6E6E6', 'CFCFCF', 'BDBDBD');   // roof vents
  isoBox(s, 3.40, 1.94, 0.16, 0.16, 0.18, 'E6E6E6', 'CFCFCF', 'BDBDBD');
  isoBox(s, 2.72, 2.28, 0.16, 0.16, 0.18, 'E6E6E6', 'CFCFCF', 'BDBDBD');
  poly(s, '5DD75D', [{ x: 1.30, y: 3.03 }, { x: 2.72, y: 3.84 }, { x: 2.72, y: 4.02 }, { x: 1.30, y: 3.21 }]);
  poly(s, '5DD75D', [{ x: 2.74, y: 3.84 }, { x: 4.14, y: 3.03 }, { x: 4.14, y: 3.21 }, { x: 2.74, y: 4.02 }]);

  // front hall ---------------------------------------------------------------
  isoBox(s, 4.51, 3.24, 1.83, 1.83, 1.30, '5DD75D', '228B22', '228B22');
  isoBox(s, 4.51, 3.14, 1.74, 1.74, 1.22, 'E3E3E3', 'C9C9C9', 'B8B8B8');
  poly(s, '080808', [{ x: 3.94, y: 4.55 }, { x: 4.60, y: 4.93 }, { x: 4.60, y: 6.05 }, { x: 3.94, y: 5.67 }]);
  poly(s, '228B22', [{ x: 4.05, y: 4.68 }, { x: 4.50, y: 4.94 }, { x: 4.50, y: 5.92 }, { x: 4.05, y: 5.66 }]);
  poly(s, '5DD75D', [{ x: 3.02, y: 4.28 }, { x: 4.52, y: 5.14 }, { x: 4.52, y: 5.32 }, { x: 3.02, y: 4.46 }]);
  poly(s, '5DD75D', [{ x: 4.54, y: 5.14 }, { x: 6.00, y: 4.30 }, { x: 6.00, y: 4.48 }, { x: 4.54, y: 5.32 }]);

  // pallet of cartons --------------------------------------------------------
  isoBox(s, 1.18, 4.06, 0.55, 0.55, 0.42, 'FFE1A3', 'EDC989', 'E0B171');
  isoBox(s, 1.18, 4.50, 0.55, 0.55, 0.42, 'FFE1A3', 'EDC989', 'E0B171');
  isoBox(s, 1.18, 4.94, 0.60, 0.60, 0.12, 'BAA366', 'B98347', 'B98347');

  // forklift -----------------------------------------------------------------
  isoBox(s, 2.28, 4.66, 0.34, 0.34, 0.52, 'FA7C20', 'D9640F', 'C25A0E');
  poly(s, '3B4042', [{ x: 2.10, y: 4.82 }, { x: 2.24, y: 4.90 }, { x: 2.24, y: 5.44 }, { x: 2.10, y: 5.36 }]);
  poly(s, 'FFE1A3', [{ x: 2.06, y: 5.02 }, { x: 2.32, y: 5.16 }, { x: 2.32, y: 5.42 }, { x: 2.06, y: 5.28 }]);
  box(s, 'ellipse', 2.27, 5.26, 0.17, 0.17, '1C1E1F');
  box(s, 'ellipse', 2.50, 5.34, 0.17, 0.17, '1C1E1F');

  // delivery truck -----------------------------------------------------------
  isoBox(s, 3.86, 5.06, 0.42, 0.66, 0.50, 'E6E6E6', 'C9C9C9', 'DADADA');   // box body
  isoBox(s, 3.55, 5.60, 0.30, 0.34, 0.40, '228B22', '196819', '1F7A1F');   // cab
  poly(s, '5DD75D', [{ x: 3.52, y: 5.98 }, { x: 3.82, y: 6.14 }, { x: 3.82, y: 6.28 }, { x: 3.52, y: 6.12 }]);
  box(s, 'ellipse', 3.72, 6.10, 0.19, 0.19, '2A2A2A');
  box(s, 'ellipse', 4.05, 6.02, 0.19, 0.19, '2A2A2A');
  return s;
}

/** Slide 25 — stylised airliner banking across the orange sun disc. */
function airplane(s) {
  poly(s, 'DCDCDC', [{ x: 10.43, y: 2.71 }, { x: 11.20, y: 3.20 }, { x: 11.10, y: 4.30 }, { x: 10.55, y: 4.10 }]); // tail fin
  poly(s, 'B8B8B8', [{ x: 11.50, y: 2.97 }, { x: 12.15, y: 3.30 }, { x: 11.60, y: 4.79 }, { x: 11.30, y: 4.30 }]); // rear stabiliser
  poly(s, 'DCDCDC', [{ x: 8.32, y: 2.42 }, { x: 8.83, y: 2.60 }, { x: 9.10, y: 3.39 }, { x: 8.60, y: 3.30 }]);     // upper wing tip
  poly(s, 'DCDCDC', [{ x: 7.47, y: 3.33 }, { x: 8.67, y: 3.36 }, { x: 9.40, y: 3.90 }, { x: 8.00, y: 3.75 }]);     // upper wing
  poly(s, 'DCDCDC', [{ x: 7.55, y: 4.48 }, { x: 9.60, y: 4.20 }, { x: 10.83, y: 5.10 }, { x: 8.00, y: 5.43 }]);    // near wing
  poly(s, 'B8B8B8', [{ x: 7.98, y: 5.08 }, { x: 10.83, y: 5.02 }, { x: 10.83, y: 5.30 }, { x: 7.98, y: 5.54 }]);   // wing shadow
  box(s, 'ellipse', 11.21, 4.44, 0.62, 0.60, 'FFC966');                                                            // nose cone
  box(s, 'roundRect', 8.05, 3.05, 3.85, 1.42, 'FBFCFE', { r: 0.7, rot: 26 });                                      // fuselage
  box(s, 'ellipse', 11.47, 3.74, 0.47, 0.47, 'EA3017');                                                            // engine (far)
  box(s, 'ellipse', 9.06, 5.02, 0.50, 0.50, 'EA3017');                                                             // engine (near)
  [0, 1, 2, 3, 4, 5, 6].forEach(function (i) {                                                                     // cabin windows
    box(s, 'ellipse', 9.00 + i * 0.345, 3.73 + i * 0.196, 0.23, 0.27, 'FFC966');
  });
  return s;
}

/** Slide 28 — four-petal marketing flower diagram. */
function petals(s) {
  const cx = 6.683, cy = 4.351, ring = 1.151, core = 0.71;
  const spokes = [[cx, cy - 1.386], [cx, cy + 1.386], [cx - 1.240, cy], [cx + 1.243, cy]];
  [[cx - 0.99, cy - 0.92], [cx + 1.04, cy - 0.93], [cx - 0.99, cy + 0.94], [cx + 1.05, cy + 0.96]]
    .forEach(function (p) { box(s, 'ellipse', p[0] - 0.75, p[1] - 0.75, 1.5, 1.5, WHITE); });
  spokes.forEach(function (p) { box(s, 'ellipse', p[0] - ring, p[1] - ring, ring * 2, ring * 2, ORANGE); });
  spokes.forEach(function (p) { box(s, 'ellipse', p[0] - core, p[1] - core, core * 2, core * 2, GREEN); });
  return s;
}


// ── slide 19 table ──────────────────────────────────────────────────────────
const TRACK_ROWS = [
  ['Metric',                '2023',   '2024',   '2025'  ],
  ['On-Time Deliveries',    '88%',    '90%',    '92%'   ],
  ['Customer Satisfaction', '89%',    '91%',    '94%'   ],
  ['Total Shipments',       '50.000', '58.000', '66.000'],
  ['Cost Efficiency',       '75%',    '78%',    '80%'   ],
];

function trackTable(s) {
  const rule = { color: ORANGE, pt: 0.5 };
  const rows = TRACK_ROWS.map(function (cells, r) {
    return cells.map(function (text, c) {
      return {
        text: text,
        options: {
          color: r === 0 ? ORANGE : GRAY_D,
          fontSize: r === 0 ? (c === 0 ? 20 : 18) : (c === 0 ? 18 : 16),
          fontFace: HEAD, valign: 'middle', align: 'left',
          margin: [0, 0, 0, 14.4],
          border: [
            { type: r === 0 ? 'none' : 'solid', color: ORANGE, pt: 0.5 },
            { type: c === 3 ? 'none' : 'solid', color: ORANGE, pt: 0.5 },
            { type: r === 4 ? 'none' : 'solid', color: ORANGE, pt: 0.5 },
            { type: c === 0 ? 'none' : 'solid', color: ORANGE, pt: 0.5 },
          ],
          fill: { color: WHITE },
        },
      };
    });
  });
  s.addTable(rows, { x: 1.382, y: 2.623, w: 10.583, colW: [3.681, 2.301, 2.301, 2.301], rowH: 0.753 });
  return s;
}

// ── slide 18 stacked bar chart ──────────────────────────────────────────────
const CHART_CATS = ['Data 1', 'Data 2', 'Data 3', 'Data 4'];
const CHART_SERIES = [
  { name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5], color: GREEN,  label: WHITE },
  { name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8], color: ORANGE, label: INK   },
  { name: 'Series 3', values: [2.0, 2.0, 3.0, 5.0], color: CREAM,  label: INK   },
];

function metricsChart(s) {
  s.addChart('bar', CHART_SERIES.map(function (ser) {
    return { name: ser.name, labels: CHART_CATS, values: ser.values };
  }), {
    x: 1.021, y: 0.991, w: 5.639, h: 5.519,
    barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 150, barOverlapPct: 100,
    chartColors: CHART_SERIES.map(function (ser) { return ser.color; }),
    showTitle: true, title: 'Logistics Efficiency Metrics',
    titleFontFace: BODY, titleFontSize: 18.6, titleColor: '262626',
    showLegend: false, showValue: true,
    dataLabelFontFace: HEAD, dataLabelFontSize: 12, dataLabelColor: GRAY_D, dataLabelPosition: 'ctr',
    catAxisLabelFontFace: BODY, catAxisLabelFontSize: 12, catAxisLabelColor: GRAY,
    valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12, valAxisLabelColor: GRAY,
    catAxisLineColor: 'D9D9D9', valAxisLineShow: false,
    valGridLine: { style: 'solid', size: 0.75, color: 'D9D9D9' },
    catGridLine: { style: 'none' },
  });
  return s;
}

// ── slides ─────────────────────────────────────────────────────────────────

/** Slide 1 — title / cover. */
function slide01(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 0, 0, 13.333, 7.5, GREEN, { a: 40 });
  box(s, 'ellipse', 10.185, 6.365, 2.827, 2.827, ORANGE, { a: 20 });
  box(s, 'ellipse', 2.604, 2.013, 2.827, 2.827, ORANGE, { a: 20 });
  tx(s, 'Logistic Presentation Template', 4.24, 3.82, 4.853, 0.438, { al: 'center', sz: 20, c: CREAM });
  tx(s, [run('Green ', { sz: 60, hd: true, c: CREAM }), run('Logistic', { sz: 66, hd: true, c: CREAM })], 2.099, 2.709, 9.135, 1.212, { al: 'center' });
  tx(s, 'Green Logistic', 9.725, 6.896, 2.966, 0.404, { al: 'right', hd: true, c: CREAM });
  box(s, 'chevron', 8.773, 3.905, 0.268, 0.268, ORANGE);
  box(s, 'chevron', 8.964, 3.905, 0.268, 0.268, CREAM);
  chrome(s, 1);
}

/** Slide 2 — intro with photo. */
function slide02(pres) {
  const s = newSlide(pres);
  box(s, 'roundRect', 3.708, -0.253, 0.597, 0.878, GREEN, { r: 0.298 });
  box(s, 'roundRect', 9.514, 4.692, 4.319, 0.986, ORANGE, { r: 0.493 });
  tx(s, 'Explore the best logistics!', 0.652, 1.945, 5.729, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, 'Logistic', 7.79, 4.694, 1.724, 1.555, { al: 'right', ls: 0.9, sb: 12, sz: 48, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  tx(s, L6, 0.652, 3.478, 6.5, 1.122, { ls: 1.5, sz: 14, c: GRAY_M });
  box(s, 'roundRect', 0.724, 4.905, 1.972, 0.52, GREEN, { r: 0.26, sh: SH_SOFT });
  tx(s, 'Get\u2019s Started!', 0.829, 4.923, 1.762, 0.415, { al: 'center', va: 'middle', ls: 1.5, sz: 14, hd: true, c: CREAM });
  box(s, 'chevron', 2.704, 4.959, 0.411, 0.411, ORANGE);
  box(s, 'chevron', 2.997, 4.959, 0.411, 0.411, CREAM);
  chrome(s, 2);
}

/** Slide 3 — section divider. */
function slide03(pres) {
  const s = newSlide(pres, GREEN_D);
  tx(s, 'Logistic Presentation Template', 4.24, 4.356, 4.853, 0.438, { al: 'center', sz: 20, c: CREAM });
  tx(s, [run('Green ', { sz: 60, hd: true, c: ORANGE }), run('Logistic', { sz: 66, hd: true, c: ORANGE })], 2.099, 3.144, 9.135, 1.212, { al: 'center' });
  box(s, 'chevron', 6.24, 2.321, 0.498, 0.498, ORANGE);
  box(s, 'chevron', 6.595, 2.321, 0.498, 0.498, CREAM);
  chrome(s, 3);
}

/** Slide 4 — driving efficiency. */
function slide04(pres) {
  const s = newSlide(pres);
  box(s, 'roundRect', 7.528, 3.708, 0.584, 0.583, GREEN_D, { r: 0.291 });
  tx(s, 'Driving Efficiency in Supply Chain', 7.448, 1.58, 4.913, 1.919, { ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, L2, 7.448, 4.319, 4.288, 1.122, { ls: 1.5, sz: 14, c: GRAY });
  box(s, 'roundRect', -0.323, 1.125, 3.74, 5.25, ORANGE, { r: 0.318 });
  box(s, 'rect', 7.406, 3.708, 0.339, 0.339, WHITE, { a: 100 });
  icon(s, 7.683, 3.852, 0.285, 0.283, CREAM);
  box(s, 'roundRect', 12.691, 1.125, 1.35, 5.25, CREAM, { r: 0.24 });
  tx(s, 'Logistic', 12.224, 1.449, 1.202, 4.601, { al: 'center', vert: 'vert270', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'roundRect', 3.764, 5.611, 3.017, 1.119, GREEN, { r: 0.206, sh: SH_SOFT });
  tx(s, '“Lorem ipsum dolor sit amet, consectetuer adipiscing elit”', 3.878, 5.773, 2.789, 0.769, { al: 'center', ls: 1.5, sz: 14, i: true, c: CREAM });
  box(s, 'chevron', 3.208, 5.855, 0.391, 0.391, ORANGE);
  box(s, 'chevron', 3.487, 5.855, 0.391, 0.391, CREAM);
  box(s, 'roundRect', 6.368, -0.2, 0.597, 0.878, GREEN, { r: 0.298 });
  chrome(s, 4);
}

/** Slide 5 — safe reliable express. */
function slide05(pres) {
  const s = newSlide(pres);
  tx(s, 'Logistic', 8.299, 6.155, 5.034, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  tx(s, 'Logistic', 0.273, 3.423, 5.034, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'roundRect', -0.509, 4.635, 14.352, 1.375, ORANGE, { r: 0.211, sh: SH_CARD });
  tx(s, 'Safe, Reliable And Express Logistic And Transport Solutions', 2.596, 1.432, 8.14, 2.121, { al: 'center', sz: 40, b: true, hd: true });
  box(s, 'roundRect', 1.506, 4.038, 4.264, 2.569, GREEN, { r: 0.395, sh: SH_SOFT });
  tx(s, L4, 2.456, 5.078, 3.251, 1.122, { ls: 1.5, sz: 14, c: WHITE });
  tx(s, 'Almost before we knew', 2.456, 4.437, 2.062, 0.64, { sz: 16, hd: true, c: CREAM });
  box(s, 'roundRect', 1.852, 4.503, 0.508, 0.508, CREAM, { r: 0.254 });
  box(s, 'rect', 1.933, 4.577, 0.346, 0.346, WHITE, { a: 100 });
  icon(s, 1.962, 4.606, 0.293, 0.293, GREEN);
  box(s, 'roundRect', 7.563, 4.028, 4.264, 2.569, GREEN, { r: 0.395, sh: SH_SOFT });
  tx(s, L4, 8.549, 5.068, 3.143, 1.122, { ls: 1.5, sz: 14, c: WHITE });
  tx(s, 'Almost before we knew it, we had left', 8.549, 4.427, 2.586, 0.64, { sz: 16, hd: true, c: CREAM });
  box(s, 'roundRect', 7.927, 4.493, 0.508, 0.508, CREAM, { r: 0.254 });
  icon(s, 8.042, 4.597, 0.135, 0.225, GREEN);
  icon(s, 8.199, 4.799, 0.121, 0.052, GREEN);
  icon(s, 8.126, 4.825, 0.079, 0.079, GREEN);
  icon(s, 8.146, 4.845, 0.039, 0.039, GREEN);
  icon(s, 8.155, 4.661, 0.161, 0.158, GREEN);
  icon(s, 8.126, 4.59, 0.104, 0.101, GREEN);
  box(s, 'chevron', 6.24, 0.655, 0.498, 0.498, ORANGE);
  box(s, 'chevron', 6.595, 0.655, 0.498, 0.498, CREAM);
  box(s, 'roundRect', 12.41, -0.222, 0.597, 0.878, CREAM, { r: 0.298 });
  chrome(s, 5);
}

/** Slide 6 — fast delivery. */
function slide06(pres) {
  const s = newSlide(pres);
  tx(s, 'Logistic', 7.993, 0.757, 2.201, 2.1, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'roundRect', 8.041, 0, 1.518, 7.5, ORANGE, { a: 22, r: 0.594 });
  box(s, 'roundRect', 7.594, 0, 1.518, 7.5, GREEN, { a: 22, r: 0.594 });
  box(s, 'roundRect', 7.087, 0, 1.518, 7.5, ORANGE, { a: 22, r: 0.594 });
  box(s, 'roundRect', 6.503, 1.629, 6.111, 3.758, GREEN, { r: 0.577, sh: SH_SOFT });
  tx(s, 'Fast Delivery, Endless Smiles', 6.992, 1.986, 4.913, 1.447, { sz: 40, hd: true, c: CREAM });
  tx(s, L3.trim(), 6.992, 4.04, 4.566, 0.769, { ls: 1.5, sz: 14, c: WHITE });
  box(s, 'roundRect', 7.147, 3.483, 0.508, 0.508, CREAM, { r: 0.254 });
  icon(s, 7.212, 3.631, 0.378, 0.211, GREEN);
  box(s, 'chevron', 11.229, 1.373, 0.498, 0.498, ORANGE);
  box(s, 'chevron', 11.584, 1.373, 0.498, 0.498, CREAM);
  box(s, 'roundRect', 12.426, -0.195, 0.597, 0.878, GREEN, { r: 0.298 });
  chrome(s, 6);
}

/** Slide 7 — global provider. */
function slide07(pres) {
  const s = newSlide(pres);
  tx(s, 'Being A Global Logistics Service Provider', 0.798, 1.831, 5.549, 1.919, { ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, L6.trim(), 0.798, 3.971, 4.632, 1.475, { ls: 1.5, sz: 14, c: GRAY });
  tx(s, 'Logistic', 7.431, 6.95, 5.054, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'roundRect', 7.391, 3.793, 0.826, 1.32, ORANGE, { r: 0.192, sh: SH_SOFT });
  box(s, 'roundRect', 9.736, 2.004, 2.376, 1.205, GREEN, { r: 0.28, sh: SH_SOFT });
  tx(s, 'Seamless Global Reach', 10.518, 2.295, 1.507, 0.572, { sz: 14, c: CREAM });
  box(s, 'roundRect', 10.013, 2.354, 0.505, 0.505, CREAM, { r: 0.253 });
  icon(s, 10.078, 2.502, 0.376, 0.21, GREEN);
  box(s, 'chevron', 11.715, 1.775, 0.435, 0.435, ORANGE);
  box(s, 'chevron', 12.025, 1.775, 0.435, 0.435, CREAM);
  box(s, 'roundRect', 12.43, -0.214, 0.597, 0.878, GREEN, { r: 0.298 });
  chrome(s, 7);
}

/** Slide 8 — air freight products. */
function slide08(pres) {
  const s = newSlide(pres);
  tx(s, 'Logistic', 11.073, 4.626, 2.201, 2.349, { al: 'center', vert: 'vert270', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  tx(s, 'General Air Freight Products', 1.352, 1.198, 5.654, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true });
  box(s, 'roundRect', 8.097, 5.247, 3.677, 1.161, ORANGE, { r: 0.326, sh: SH_SOFT });
  box(s, 'roundRect', 0.545, 2.773, 3.27, 1.703, WHITE, { r: 0.479, sh: SH_SOFT });
  tx(s, L5, 1.352, 3.422, 2.443, 0.769, { ls: 1.5, sz: 14, c: GRAY_D });
  tx(s, 'Standard Shipping', 1.352, 3.052, 3.017, 0.37, { sz: 16, b: true, c: ORANGE });
  box(s, 'roundRect', 0.747, 2.994, 0.508, 0.508, GREEN, { r: 0.254 });
  icon(s, 0.842, 3.116, 0.318, 0.263, WHITE);
  box(s, 'roundRect', 4.011, 3.203, 3.27, 1.703, WHITE, { r: 0.479, sh: SH_SOFT });
  tx(s, L5, 4.891, 3.858, 2.357, 0.769, { ls: 1.5, sz: 14, c: GRAY_D });
  tx(s, 'Express Delivery', 4.891, 3.51, 3.017, 0.37, { sz: 16, b: true, c: ORANGE });
  box(s, 'roundRect', 4.287, 3.43, 0.508, 0.508, GREEN, { r: 0.254 });
  icon(s, 4.352, 3.578, 0.378, 0.211, WHITE);
  box(s, 'roundRect', 1.75, 5.054, 3.27, 1.703, WHITE, { r: 0.479, sh: SH_SOFT });
  box(s, 'roundRect', 2.063, 5.279, 0.485, 0.508, GREEN, { r: 0.242 });
  tx(s, L5, 2.641, 5.75, 2.252, 0.769, { ls: 1.5, sz: 14, c: GRAY_D });
  tx(s, 'Special Cargo', 2.641, 5.379, 2.883, 0.37, { sz: 16, b: true, c: ORANGE });
  icon(s, 2.155, 5.38, 0.301, 0.305, WHITE);
  box(s, 'chevron', 7.399, 5.176, 0.407, 0.407, ORANGE);
  box(s, 'chevron', 7.689, 5.176, 0.407, 0.407, CREAM);
  box(s, 'roundRect', 11.704, -0.148, 0.552, 0.878, GREEN, { r: 0.276 });
  chrome(s, 8);
}

/** Slide 9 — fast and reliable shipping. */
function slide09(pres) {
  const s = newSlide(pres, GREEN);
  tx(s, 'Logistic', 8.558, 5.372, 5.034, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Logistic', 5.855, 3.123, 5.034, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Logistic', 2.648, 4.914, 5.034, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Logistic', -0.241, 3.415, 5.034, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Fast and Reliable Shipping', 2.484, 1.251, 8.507, 0.774, { al: 'center', sz: 40, hd: true, c: WHITE });
  box(s, 'roundRect', 0.91, 2.425, 3.316, 4.183, ORANGE, { r: 0.336, sh: SH_SOFT });
  box(s, 'roundRect', 4.883, 2.425, 3.316, 4.183, ORANGE, { r: 0.315, sh: SH_SOFT });
  box(s, 'roundRect', 8.856, 2.425, 3.316, 4.183, ORANGE, { r: 0.315, sh: SH_SOFT });
  box(s, 'chevron', 6.24, 0.595, 0.498, 0.498, ORANGE);
  box(s, 'chevron', 6.595, 0.595, 0.498, 0.498, CREAM);
  box(s, 'roundRect', 12.302, -0.221, 0.597, 0.878, ORANGE, { r: 0.298 });
  chrome(s, 9);
}

/** Slide 10 — proud to deliver. */
function slide10(pres) {
  const s = newSlide(pres);
  tx(s, 'Logistic', 8.373, 1.13, 4.551, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  tx(s, 'Logistic', 8.481, 5.366, 5.034, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Proud to Deliver Excellence', 0.778, 1.819, 6.667, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true });
  box(s, 'roundRect', 11.653, 1.653, 0.903, 4.167, GREEN, { r: 0.368 });
  tx(s, L7, 0.778, 4.678, 4.91, 0.769, { ls: 1.5, sz: 14, c: GRAY });
  box(s, 'roundRect', 0.922, 3.619, 0.563, 0.589, GREEN, { r: 0.281 });
  icon(s, 0.987, 3.697, 0.433, 0.433, CREAM);
  tx(s, 'Excellence in Every Step', 0.778, 4.292, 3.673, 0.404, { hd: true, c: ORANGE });
  box(s, 'chevron', 6.542, 1.497, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 6.804, 1.497, 0.368, 0.368, CREAM);
  box(s, 'roundRect', 11.805, -0.189, 0.597, 0.878, CREAM, { r: 0.298 });
  chrome(s, 10);
}

/** Slide 11 — complete logistics services. */
function slide11(pres) {
  const s = newSlide(pres);
  tx(s, 'Logistic', 5.553, 6.283, 4.551, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'roundRect', 0, 0, 1.143, 7.5, ORANGE, { r: 0 });
  tx(s, 'Complete Logistics Services', 7.391, 2.194, 5.413, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, L2 + ' ', 7.391, 3.764, 4.29, 1.122, { ls: 1.5, sz: 14, c: GRAY });
  box(s, 'chevron', 7.474, 1.612, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 7.736, 1.612, 0.368, 0.368, CREAM);
  box(s, 'roundRect', 12.677, -0.177, 0.597, 0.878, GREEN, { r: 0.298 });
  box(s, 'roundRect', 6.064, 5.604, 2.376, 1.122, GREEN, { r: 0.244, sh: SH_SOFT });
  tx(s, 'End-to-End Solutions', 6.885, 5.879, 1.507, 0.572, { sz: 14, c: CREAM });
  box(s, 'roundRect', 6.341, 5.913, 0.505, 0.505, CREAM, { r: 0.253 });
  icon(s, 6.459, 6.017, 0.269, 0.296, GREEN);
  chrome(s, 11);
}

/** Slide 12 — explore the logistic service. */
function slide12(pres) {
  const s = newSlide(pres);
  box(s, 'ellipse', -1.255, 1.4, 7.221, 7.221, ORANGE);
  box(s, 'rect', 7.797, 0, 5.036, 7.5, GREEN);
  tx(s, 'Explore the logistic service', 0.578, 2.762, 5.514, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true, c: WHITE });
  tx(s, L2, 0.578, 4.228, 4.944, 1.122, { ls: 1.5, sz: 14, c: WHITE });
  box(s, 'roundRect', 6.916, 0.936, 4.302, 1.703, WHITE, { r: 0.479, sh: SH_SOFT });
  tx(s, L1 + ' ', 7.723, 1.585, 3.178, 0.769, { ls: 1.5, sz: 14, c: GRAY });
  tx(s, 'Comprehensive Solutions', 7.723, 1.215, 3.237, 0.37, { sz: 16, hd: true, c: ORANGE });
  box(s, 'roundRect', 7.136, 1.215, 0.508, 0.508, GREEN, { r: 0.254 });
  box(s, 'roundRect', 6.916, 2.972, 4.302, 1.703, WHITE, { r: 0.479, sh: SH_SOFT });
  tx(s, L1 + ' ', 7.797, 3.627, 2.991, 0.769, { ls: 1.5, sz: 14, c: GRAY });
  box(s, 'roundRect', 7.192, 3.199, 0.508, 0.508, GREEN, { r: 0.254 });
  box(s, 'roundRect', 6.916, 5.009, 4.302, 1.703, WHITE, { r: 0.479, sh: SH_SOFT });
  box(s, 'roundRect', 7.23, 5.233, 0.485, 0.508, GREEN, { r: 0.242 });
  tx(s, L1 + ' ', 7.808, 5.704, 3.152, 0.769, { ls: 1.5, sz: 14, c: GRAY });
  tx(s, 'Expert Support', 7.808, 5.334, 3.093, 0.37, { sz: 16, hd: true, c: ORANGE });
  tx(s, 'Real-Time Tracking', 7.808, 3.253, 3.093, 0.37, { sz: 16, hd: true, c: ORANGE });
  icon(s, 7.24, 1.316, 0.301, 0.305, CREAM);
  icon(s, 7.272, 5.29, 0.415, 0.415, CREAM);
  icon(s, 7.265, 3.294, 0.358, 0.358, CREAM);
  box(s, 'roundRect', 12.677, 1.778, 1.35, 4.934, CREAM, { r: 0.24 });
  tx(s, 'Logistic', 11.759, 2.279, 1.202, 3.93, { al: 'center', vert: 'vert270', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  box(s, 'chevron', 5.049, 0.936, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 5.311, 0.936, 0.368, 0.368, CREAM);
  box(s, 'roundRect', 0.69, 5.459, 1.972, 0.446, GREEN, { r: 0.223, sh: SH_SOFT });
  tx(s, 'Get\u2019s Started!', 0.795, 5.58, 1.762, 0.144, { al: 'center', va: 'middle', ls: 1.5, sz: 14, hd: true, c: CREAM });
  box(s, 'roundRect', 12.677, -0.177, 0.597, 0.878, CREAM, { r: 0.298 });
  chrome(s, 12);
}

/** Slide 13 — contract logistics provider. */
function slide13(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 10.084, 4.799, 3.289, 2.701, ORANGE);
  tx(s, 'Logistic', 10.42, 5.157, 2.617, 2.1, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'rect', 0, 4.799, 3.289, 2.701, GREEN);
  tx(s, 'Logistic', 0.294, 5.157, 2.617, 2.1, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  box(s, 'roundRect', 1.201, 3.931, 2.743, 2.219, WHITE, { r: 0.289, sh: SH_CARD });
  tx(s, [run('World\u2019s leading Contract ', { brk: true }), run('Logistics ProviderA')], 2.514, 1.125, 8.319, 1.313, { al: 'center', ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, 'Fast & Good Response', 1.612, 5.04, 1.921, 0.707, { al: 'center', hd: true, c: ORANGE });
  box(s, 'roundRect', 3.264, 4.1, 0.508, 0.508, ORANGE, { r: 0.254 });
  tx(s, '1.', 3.289, 4.152, 0.457, 0.404, { al: 'center', hd: true, c: WHITE });
  box(s, 'roundRect', 5.282, 4.608, 2.743, 2.219, WHITE, { r: 0.289, sh: SH_CARD });
  tx(s, 'Logistic Delivery', 5.693, 5.717, 1.921, 0.707, { al: 'center', hd: true, c: GREEN });
  box(s, 'roundRect', 7.344, 4.778, 0.508, 0.508, GREEN, { r: 0.254 });
  tx(s, '2.', 7.37, 4.83, 0.457, 0.404, { al: 'center', hd: true, c: CREAM });
  box(s, 'roundRect', 9.412, 4.183, 2.743, 2.219, WHITE, { r: 0.289, sh: SH_CARD });
  tx(s, 'Fast & Efficient Delivery', 9.618, 5.293, 2.332, 0.707, { al: 'center', hd: true, c: ORANGE });
  box(s, 'roundRect', 11.475, 4.353, 0.508, 0.508, ORANGE, { r: 0.254 });
  tx(s, '3.', 11.5, 4.405, 0.457, 0.404, { al: 'center', hd: true, c: WHITE });
  box(s, 'roundRect', 12.236, -0.218, 0.597, 0.878, GREEN, { r: 0.298 });
  box(s, 'chevron', 11.098, 1.781, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 11.361, 1.781, 0.368, 0.368, CREAM);
  chrome(s, 13);
}

/** Slide 14 — shipping product safely. */
function slide14(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 2.558, 2.713, 1.373, 1.995, GREEN, { rot: -90 });
  box(s, 'rect', 9.388, 5.679, 1.373, 1.821, ORANGE);
  box(s, 'rect', 0, 0, 3.035, 7.5, GREEN);
  tx(s, 'Logistic', -0.278, 1.181, 1.202, 4.927, { al: 'center', vert: 'vert270', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Logistic', 9.403, 0.909, 4.551, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'roundRect', 7.315, 1.595, 5.518, 4.264, WHITE, { r: 0.637, sh: SH_CARD });
  tx(s, 'Shipping Product Safely', 7.703, 2.666, 4.777, 1.447, { sz: 40, hd: true });
  tx(s, L2, 7.703, 4.242, 4.319, 1.122, { ls: 1.5, sz: 14, c: GRAY });
  box(s, 'roundRect', 0.579, 2.579, 3.035, 3.279, WHITE, { r: 0.506, sh: SH_CARD });
  box(s, 'roundRect', 3.891, 1.584, 3.035, 3.279, ORANGE, { r: 0.506, sh: SH_CARD });
  tx(s, '01', 0.856, 2.719, 2.538, 0.926, { va: 'middle', ls: 1.3, sz: 40, hd: true, c: GREEN });
  tx(s, 'Secure Packaging', 0.856, 4.017, 2.827, 0.404, { hd: true, c: GREEN });
  tx(s, L1, 0.856, 4.503, 2.299, 1.002, { va: 'middle', sz: 14, c: GRAY });
  tx(s, '02', 4.25, 1.789, 2.538, 0.926, { va: 'middle', ls: 1.3, sz: 40, hd: true, c: WHITE });
  tx(s, 'Live Tracking', 4.25, 3.023, 2.827, 0.404, { hd: true, c: WHITE });
  tx(s, L1, 4.25, 3.525, 2.299, 1.002, { va: 'middle', sz: 14, c: WHITE });
  box(s, 'line', 4.396, 3.452, 2.037, 0, null, { ln: GREEN, lw: 2 });
  box(s, 'line', 0.974, 4.451, 2.026, 0, null, { ln: ORANGE, lw: 2 });
  box(s, 'roundRect', 7.443, -0.21, 0.597, 0.878, ORANGE, { r: 0.298 });
  box(s, 'chevron', 7.856, 2.086, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 8.119, 2.086, 0.368, 0.368, CREAM);
  box(s, 'roundRect', 2.758, 2.99, 0.508, 0.508, GREEN, { r: 0.254 });
  icon(s, 2.875, 3.106, 0.273, 0.277, CREAM);
  box(s, 'roundRect', 6.025, 2.071, 0.508, 0.508, WHITE, { r: 0.254 });
  icon(s, 6.098, 2.166, 0.358, 0.358, GREEN);
  chrome(s, 14);
}

/** Slide 15 — moving products 1-2. */
function slide15(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 0, 0, 6.667, 3.75, GREEN);
  box(s, 'rect', 0, 3.75, 6.667, 3.764, ORANGE);
  tx(s, [run('Moving Your Products ', { brk: true }), run('Across All Borders')], 7.12, 1.506, 5.678, 1.919, { ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, L2, 1.028, 4.906, 4.944, 1.122, { ls: 1.5, sz: 14, c: WHITE });
  tx(s, 'Logistic', -0.584, 3.024, 4.275, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, L2, 1.028, 1.722, 4.944, 1.122, { ls: 1.5, sz: 14, c: WHITE });
  tx(s, '2. Safe and Secure Delivery', 1.028, 4.482, 4.375, 0.413, { ls: 1.1, b: true, hd: true, c: GREEN });
  tx(s, '1. Road Freight Service', 1.028, 1.298, 4.375, 0.413, { hd: true, c: ORANGE });
  tx(s, 'Logistic', 4.871, 6.329, 4.275, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'chevron', 7.251, 0.972, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 7.513, 0.972, 0.368, 0.368, CREAM);
  box(s, 'roundRect', 12.236, -0.218, 0.597, 0.878, ORANGE, { r: 0.298 });
  tx(s, L2 + ' Cum sociis natoque penatibus et magnis dis.', 7.12, 4.882, 5.332, 1.122, { ls: 1.5, sz: 14, c: GRAY });
  box(s, 'roundRect', 7.248, 3.651, 0.577, 0.577, GREEN, { r: 0.288 });
  icon(s, 7.331, 3.741, 0.411, 0.411, CREAM);
  tx(s, 'Excellence in Every Delivery', 7.12, 4.3, 5.533, 0.549, { ls: 1.5, sz: 20, hd: true, c: ORANGE });
  chrome(s, 15);
}

/** Slide 16 — moving products 3-4. */
function slide16(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 0, 3.75, 6.667, 3.764, ORANGE);
  box(s, 'rect', 6.667, 3.75, 6.667, 3.75, GREEN);
  tx(s, [run('Moving Your Products ', { brk: true }), run('Across All Borders')], 1.042, 1.502, 6.806, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, L2, 1, 4.878, 4.944, 1.122, { ls: 1.5, sz: 14, c: WHITE });
  tx(s, 'Logistic', 9.059, 6.858, 4.275, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, L2, 7.514, 4.878, 4.944, 1.122, { ls: 1.5, sz: 14, c: WHITE });
  tx(s, '3. Efficient Transit Time', 1.014, 4.454, 4.375, 0.404, { hd: true, c: GREEN });
  tx(s, '4. Cost-Effective Solutions', 7.528, 4.454, 4.375, 0.413, { ls: 1.1, hd: true, c: ORANGE });
  tx(s, LQ, 8.625, 1.774, 3.694, 0.769, { ls: 1.5, sz: 14, i: true, c: GRAY });
  tx(s, 'Logistic', -1.138, 3.255, 4.275, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'roundRect', 12.236, -0.218, 0.597, 0.878, GREEN, { r: 0.298 });
  box(s, 'chevron', 6.667, 2.281, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 6.929, 2.281, 0.368, 0.368, CREAM);
  chrome(s, 16);
}

/** Slide 17 — committed to excellence (gauges). */
function slide17(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 4.077, 2.905, 9.256, 1.69, GREEN);
  box(s, 'roundRect', 7.771, 1.392, 5.062, 4.662, WHITE, { r: 0.608, sh: SH_CARD });
  box(s, 'rect', 0, 2.905, 3.103, 1.69, ORANGE);
  tx(s, 'Committed to Excellence Every Time', 8.105, 2.302, 4.745, 2.121, { sz: 40, hd: true });
  tx(s, L2, 8.105, 4.526, 4.19, 1.122, { ls: 1.5, sz: 14, c: GRAY });
  box(s, 'roundRect', 0.542, 1.445, 3.128, 4.61, GREEN, { r: 0.327, sh: SH_CARD });
  tx(s, 'Create Effortlessly', 0.866, 1.586, 2.48, 0.774, { al: 'center', sz: 20, hd: true, c: WHITE });
  box(s, 'ellipse', 1.108, 2.905, 1.996, 1.996, null, { ln: WHITE, lw: 16, la: 85 });
  arc(s, 1.108, 2.905, 1.996, 1.996, 270, 321.695, { ln: WHITE, lw: 10.5 });
  tx(s, '97B', 1.324, 3.393, 1.564, 0.909, { nowrap: true, sz: 48, hd: true, c: WHITE });
  tx(s, '/container', 1.589, 4.061, 1.034, 0.34, { al: 'center', ls: 1.3, sz: 12, c: WHITE });
  tx(s, L5, 0.832, 5.192, 2.548, 0.572, { al: 'center', sz: 14, c: WHITE });
  box(s, 'line', 0.542, 2.474, 3.128, 0, null, { ln: WHITE, la: 84, lw: 1.75 });
  box(s, 'roundRect', 4.077, 1.445, 3.128, 4.61, ORANGE, { r: 0.299, sh: SH_CARD });
  tx(s, 'Everyday Adventure', 4.179, 1.586, 2.924, 0.774, { al: 'center', sz: 20, hd: true, c: WHITE });
  box(s, 'ellipse', 4.643, 2.905, 1.996, 1.996, null, { ln: '262626', lw: 16, la: 95 });
  arc(s, 4.643, 2.905, 1.996, 1.996, 270, 232.871, { ln: GREEN, lw: 10.5 });
  tx(s, '74B', 4.842, 3.393, 1.599, 0.909, { nowrap: true, sz: 48, hd: true, c: WHITE });
  tx(s, '/container', 4.906, 4.061, 1.469, 0.34, { al: 'center', ls: 1.3, sz: 12, c: WHITE });
  tx(s, L5, 4.367, 5.192, 2.548, 0.572, { al: 'center', sz: 14, c: WHITE });
  box(s, 'line', 4.077, 2.474, 3.128, 0, null, { ln: WHITE, la: 84, lw: 1.75 });
  box(s, 'roundRect', 7.952, -0.191, 0.597, 0.878, GREEN, { r: 0.298 });
  box(s, 'chevron', 8.229, 1.806, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 8.491, 1.806, 0.368, 0.368, CREAM);
  tx(s, 'Logistic', 9.574, -0.302, 4.275, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  tx(s, 'Logistic', 6.46, 6.76, 4.275, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  chrome(s, 17);
}

/** Slide 18 — performance overview (chart). */
function slide18(pres) {
  const s = newSlide(pres, GREEN);
  tx(s, 'Logistic', 3.496, 6.242, 4.275, 1.101, { al: 'center', ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  box(s, 'rect', -0.095, 2.982, 3.103, 1.69, ORANGE);
  box(s, 'roundRect', 8.071, 2.393, 4.711, 0.966, WHITE, { r: 0.293, sh: SH_CARD });
  box(s, 'roundRect', 0.542, 0.743, 6.597, 6.014, WHITE, { r: 0.785, sh: SH_CARD });
  tx(s, 'Performance Overview', 7.919, 0.868, 4.745, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true, c: WHITE });
  metricsChart(s);
  box(s, 'roundRect', 7.817, 2.251, 0.508, 0.508, ORANGE, { r: 0.254 });
  tx(s, '1.', 7.843, 2.303, 0.457, 0.404, { al: 'center', hd: true, c: WHITE });
  box(s, 'roundRect', 8.097, 3.515, 4.711, 0.966, WHITE, { r: 0.293, sh: SH_CARD });
  box(s, 'roundRect', 7.843, 3.373, 0.508, 0.508, ORANGE, { r: 0.254 });
  tx(s, '2.', 7.868, 3.425, 0.457, 0.404, { al: 'center', hd: true, c: WHITE });
  box(s, 'roundRect', 8.122, 4.637, 4.711, 0.966, WHITE, { r: 0.287, sh: SH_CARD });
  box(s, 'roundRect', 7.868, 4.495, 0.508, 0.508, ORANGE, { r: 0.254 });
  tx(s, '3.', 7.894, 4.547, 0.457, 0.404, { al: 'center', hd: true, c: WHITE });
  box(s, 'roundRect', 8.148, 5.759, 4.711, 0.966, WHITE, { r: 0.33, sh: SH_CARD });
  box(s, 'roundRect', 7.894, 5.617, 0.508, 0.508, ORANGE, { r: 0.254 });
  tx(s, '4.', 7.919, 5.669, 0.457, 0.404, { al: 'center', hd: true, c: WHITE });
  tx(s, [run('On-Time Deliveries:', { b: true, hd: true }), run(' 85%', { hd: true })], 8.919, 2.67, 3.567, 0.413, { ls: 1.1 });
  icon(s, 8.508, 2.71, 0.353, 0.332, GREEN);
  tx(s, [run('Customer Satisfaction:', { b: true, hd: true }), run(' 92%', { hd: true })], 8.919, 3.792, 3.801, 0.413, { ls: 1.1 });
  tx(s, [run('Cost Efficiency:', { b: true, hd: true }), run(' 78%', { hd: true })], 8.919, 4.887, 3.567, 0.413, { ls: 1.1 });
  tx(s, [run('Operational Uptime:', { b: true, hd: true }), run(' 95%', { hd: true })], 8.919, 6.002, 3.567, 0.413, { ls: 1.1 });
  icon(s, 8.515, 3.827, 0.342, 0.342, GREEN);
  icon(s, 8.506, 6.045, 0.344, 0.327, GREEN);
  icon(s, 8.503, 4.919, 0.35, 0.35, GREEN);
  box(s, 'roundRect', 12.768, -0.187, 0.597, 0.878, ORANGE, { r: 0.298 });
  box(s, 'chevron', 6.873, 0.992, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 7.135, 0.992, 0.368, 0.368, CREAM);
  chrome(s, 18);
}

/** Slide 19 — our track record (table). */
function slide19(pres) {
  const s = newSlide(pres, GREEN);
  box(s, 'rect', 5.722, 5.863, 1.819, 1.637, ORANGE);
  box(s, 'roundRect', 1.014, 2.306, 11.319, 4.403, WHITE, { r: 0.734, sh: SH_CARD });
  trackTable(s);
  tx(s, 'Our Track Record', 1.382, 1.057, 6.16, 0.774, { sz: 40, hd: true, c: WHITE });
  box(s, 'roundRect', 12.466, -0.175, 0.597, 0.878, ORANGE, { r: 0.298 });
  tx(s, LQ, 8.451, 1.061, 3.694, 0.769, { ls: 1.5, sz: 14, i: true, c: WHITE });
  box(s, 'chevron', 11.226, 1.322, 0.368, 0.368, ORANGE);
  box(s, 'chevron', 11.488, 1.322, 0.368, 0.368, CREAM);
  chrome(s, 19);
}

/** Slide 20 — our expert team. */
function slide20(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 0, 3.774, 13.333, 3.726, GREEN);
  tx(s, 'Logistic', 8.047, 5.383, 4.677, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Logistic', 5.646, -0.203, 4.438, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'ellipse', -0.651, 1.018, 5.416, 5.416, ORANGE);
  tx(s, L3, 0.557, 3.837, 3.758, 1.122, { ls: 1.5, sz: 14, c: WHITE });
  tx(s, 'Our Expert Team', 0.557, 2.278, 4.306, 1.447, { sz: 40, hd: true, c: WHITE });
  box(s, 'roundRect', 5.345, 5.135, 2.123, 1.406, WHITE, { r: 0.298, sh: SH_TIGHT });
  box(s, 'roundRect', 7.785, 3.781, 2.145, 1.406, WHITE, { r: 0.298, sh: SH_TIGHT });
  box(s, 'roundRect', 10.247, 4.365, 2.123, 1.406, WHITE, { r: 0.298, sh: SH_TIGHT });
  box(s, 'roundRect', 12.236, -0.155, 0.597, 0.878, ORANGE, { r: 0.298 });
  box(s, 'chevron', 5.042, 1.295, 0.498, 0.498, ORANGE);
  box(s, 'chevron', 5.398, 1.295, 0.498, 0.498, CREAM);
  tx(s, 'Francois Mercer', 5.589, 5.343, 1.634, 0.64, { al: 'center', sz: 16, hd: true, c: ORANGE });
  tx(s, 'Job Title', 5.577, 5.989, 1.634, 0.337, { al: 'center', sz: 14, i: true, hd: true, c: GRAY });
  tx(s, 'Helene Paquet', 8.047, 3.989, 1.634, 0.64, { al: 'center', sz: 16, hd: true, c: ORANGE });
  tx(s, 'Job Title', 8.035, 4.635, 1.634, 0.337, { al: 'center', sz: 14, i: true, hd: true, c: GRAY });
  tx(s, 'Benjamin Shah', 10.489, 4.572, 1.634, 0.64, { al: 'center', sz: 16, hd: true, c: ORANGE });
  tx(s, 'Job Title', 10.477, 5.218, 1.634, 0.337, { al: 'center', sz: 14, i: true, hd: true, c: GRAY });
  chrome(s, 20);
}

/** Slide 21 — our amazing team. */
function slide21(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 6.667, 3.747, 6.667, 3.753, GREEN);
  box(s, 'rect', -0.002, 3.747, 6.667, 3.753, ORANGE);
  tx(s, 'Logistic', -0.197, 1.979, 2.517, 2.1, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  tx(s, 'Logistic', 7.674, 5.319, 4.677, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Our Amazing Team', 3.575, 0.934, 6.183, 0.774, { sz: 40, hd: true });
  box(s, 'roundRect', 1.631, 2.181, 3.464, 0.663, GREEN, { r: 0.332, sh: SH_TIGHT });
  tx(s, 'Marketing Team', 1.93, 2.294, 2.866, 0.438, { al: 'center', sz: 20, hd: true, c: CREAM });
  box(s, 'roundRect', 8.285, 2.181, 3.464, 0.663, ORANGE, { r: 0.332, sh: SH_TIGHT });
  tx(s, 'Strategy Team', 8.584, 2.294, 2.866, 0.438, { al: 'center', sz: 20, hd: true, c: WHITE });
  box(s, 'roundRect', 0.686, 5.817, 1.634, 0.787, WHITE, { r: 0.172, sh: SH_TIGHT });
  tx(s, 'Chiaki Sato', 0.799, 5.891, 1.408, 0.64, { al: 'center', sz: 16, hd: true, c: GREEN });
  box(s, 'roundRect', 2.545, 5.129, 1.634, 0.787, WHITE, { r: 0.172, sh: SH_TIGHT });
  tx(s, 'Aaron Loeb', 2.63, 5.203, 1.465, 0.64, { al: 'center', sz: 16, hd: true, c: GREEN });
  box(s, 'roundRect', 4.397, 5.543, 1.634, 0.787, WHITE, { r: 0.172, sh: SH_TIGHT });
  tx(s, 'Daniel Gallego', 4.397, 5.617, 1.634, 0.64, { al: 'center', sz: 16, hd: true, c: GREEN });
  box(s, 'roundRect', 7.294, 5.423, 1.634, 0.787, WHITE, { r: 0.172, sh: SH_TIGHT });
  tx(s, 'Estelle Darcy', 7.294, 5.497, 1.634, 0.64, { al: 'center', sz: 16, hd: true, c: ORANGE });
  box(s, 'roundRect', 9.2, 4.955, 1.634, 0.787, WHITE, { r: 0.172, sh: SH_TIGHT });
  tx(s, 'Hannah Morales', 9.2, 5.029, 1.634, 0.64, { al: 'center', sz: 16, hd: true, c: ORANGE });
  box(s, 'roundRect', 11.097, 5.812, 1.634, 0.787, WHITE, { r: 0.172, sh: SH_TIGHT });
  tx(s, 'Matt Zhang', 11.144, 5.886, 1.54, 0.64, { al: 'center', sz: 16, hd: true, c: ORANGE });
  box(s, 'roundRect', 12.236, -0.155, 0.597, 0.878, ORANGE, { r: 0.298 });
  box(s, 'chevron', 6.263, 2.236, 0.498, 0.498, ORANGE);
  box(s, 'chevron', 6.619, 2.236, 0.498, 0.498, CREAM);
  chrome(s, 21);
}

/** Slide 22 — expedition progress. */
function slide22(pres) {
  const s = newSlide(pres);
  tx(s, 'Logistic', 4.057, 5.849, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  box(s, 'rect', 10.272, 3.941, 3.072, 3.559, GREEN);
  box(s, 'roundRect', 7.313, 1.786, 6.668, 3.878, ORANGE, { r: 0.534, sh: SH_SOFT });
  box(s, 'roundRect', 3.595, 4.771, 3.072, 1.504, WHITE, { r: 0.323, sh: SH_TIGHT });
  box(s, 'roundRect', 3.595, 3.071, 3.072, 1.504, WHITE, { r: 0.323, sh: SH_TIGHT });
  box(s, 'roundRect', 3.595, 1.375, 3.072, 1.504, WHITE, { r: 0.323, sh: SH_TIGHT });
  box(s, 'line', 2.77, 0.948, 0, 5.753, null, { ln: CREAM, lw: 1.5 });
  box(s, 'ellipse', 2.334, 1.592, 0.873, 0.873, WHITE, { sh: SH_SOFT });
  tx(s, 'Expedition Progress', 7.734, 2.295, 4.777, 1.447, { sz: 40, hd: true, c: WHITE });
  tx(s, L2, 7.734, 3.99, 4.319, 1.122, { ls: 1.5, sz: 14, c: WHITE });
  box(s, 'roundRect', 7.312, -0.164, 0.597, 0.878, GREEN, { r: 0.298 });
  box(s, 'ellipse', 2.334, 3.314, 0.873, 0.873, WHITE, { sh: SH_SOFT });
  box(s, 'ellipse', 2.334, 5.035, 0.873, 0.873, WHITE, { sh: SH_SOFT });
  tx(s, L1, 3.755, 1.932, 2.912, 0.686, { va: 'middle', c: GRAY });
  tx(s, 'Checkpoint 1', 3.755, 1.485, 2.912, 0.459, { va: 'middle', ls: 1.3, hd: true, c: ORANGE });
  tx(s, 'Day 7', 1.014, 1.799, 1.168, 0.459, { va: 'middle', ls: 1.3, hd: true, c: GREEN });
  tx(s, L1, 3.755, 3.608, 2.912, 0.686, { va: 'middle', c: GRAY });
  tx(s, 'Midway Review', 3.755, 3.161, 2.912, 0.459, { va: 'middle', ls: 1.3, hd: true, c: ORANGE });
  tx(s, 'Day 10', 1.014, 3.475, 1.168, 0.459, { va: 'middle', ls: 1.3, hd: true, c: GREEN });
  tx(s, L1, 3.755, 5.284, 2.912, 0.686, { va: 'middle', c: GRAY });
  tx(s, 'Final Destination', 3.755, 4.837, 2.912, 0.459, { va: 'middle', ls: 1.3, hd: true, c: ORANGE });
  tx(s, 'Day 14', 1.014, 5.151, 1.168, 0.459, { va: 'middle', ls: 1.3, hd: true, c: GREEN });
  icon(s, 2.577, 3.557, 0.385, 0.385, ORANGE);
  icon(s, 2.634, 1.841, 0.273, 0.375, ORANGE);
  icon(s, 2.543, 5.245, 0.453, 0.453, ORANGE);
  tx(s, 'Logistic', 9.813, -0.452, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'chevron', 7.406, 1.039, 0.368, 0.45, ORANGE);
  box(s, 'chevron', 7.669, 1.039, 0.368, 0.45, CREAM);
  chrome(s, 22);
}

/** Slide 23 — expedition timeline. */
function slide23(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 0, 4.819, 13.333, 2.725, GREEN, { a: 40 });
  icon(s, 8.051, 3.487, 0.629, 0.629, GREEN);
  icon(s, 4.397, 3.487, 0.629, 0.629, GREEN);
  tx(s, 'Logistic', 7.411, 6.147, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  tx(s, 'Logistic', 0.321, 2.82, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Expedition Timeline Overview', 1.014, 1.168, 7.267, 1.447, { sz: 40, hd: true });
  tx(s, LQ, 8.874, 1.507, 3.671, 0.769, { ls: 1.5, sz: 14, i: true, c: GRAY });
  box(s, 'roundRect', 1.543, 3.19, 2.984, 3.389, WHITE, { r: 0.317, sh: SH_TIGHT });
  tx(s, L1, 1.822, 5.125, 2.426, 0.992, { al: 'center', va: 'middle', c: GRAY });
  tx(s, 'Planning Stage', 1.822, 4.62, 2.426, 0.459, { al: 'center', va: 'middle', ls: 1.3, hd: true, c: ORANGE });
  box(s, 'roundRect', 5.208, 3.19, 2.984, 3.389, WHITE, { r: 0.317, sh: SH_TIGHT });
  box(s, 'roundRect', 8.874, 3.19, 2.984, 3.389, WHITE, { r: 0.317, sh: SH_TIGHT });
  tx(s, L1, 5.487, 5.125, 2.426, 0.992, { al: 'center', va: 'middle', c: GRAY });
  tx(s, 'Preparation', 5.487, 4.62, 2.426, 0.459, { al: 'center', va: 'middle', ls: 1.3, hd: true, c: ORANGE });
  tx(s, L1, 9.153, 5.125, 2.426, 0.992, { al: 'center', va: 'middle', c: GRAY });
  tx(s, 'Transport Initiation', 8.962, 4.62, 2.807, 0.459, { al: 'center', va: 'middle', ls: 1.3, hd: true, c: ORANGE });
  tx(s, 'Day 1-3', 2.451, 4.148, 1.168, 0.459, { al: 'center', va: 'middle', ls: 1.3, hd: true, c: GREEN });
  tx(s, 'Day 4-5', 5.978, 4.148, 1.445, 0.459, { al: 'center', va: 'middle', ls: 1.3, hd: true, c: GREEN });
  tx(s, 'Day 6', 9.781, 4.148, 1.168, 0.459, { al: 'center', va: 'middle', ls: 1.3, hd: true, c: GREEN });
  box(s, 'roundRect', 7.312, -0.164, 0.597, 0.878, GREEN, { r: 0.298 });
  box(s, 'chevron', 4.12, 2.029, 0.368, 0.45, ORANGE);
  box(s, 'chevron', 4.382, 2.029, 0.368, 0.45, CREAM);
  box(s, 'roundRect', 2.777, 3.577, 0.508, 0.508, GREEN, { r: 0.254 });
  box(s, 'rect', 2.882, 3.682, 0.298, 0.298, CREAM, { a: 100 });
  icon(s, 2.907, 3.695, 0.248, 0.248, CREAM);
  box(s, 'roundRect', 6.413, 3.545, 0.508, 0.508, GREEN, { r: 0.254 });
  icon(s, 6.53, 3.66, 0.273, 0.277, CREAM);
  box(s, 'roundRect', 10.048, 3.512, 0.508, 0.508, GREEN, { r: 0.254 });
  icon(s, 10.113, 3.66, 0.378, 0.211, CREAM);
  chrome(s, 23);
}

/** Slide 24 — world’s leading contract logistic. */
function slide24(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 10.167, 0, 3.167, 7.5, GREEN);
  tx(s, 'Logistic', 5.011, 1.044, 2.548, 2.1, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  box(s, 'roundRect', 6.565, 0.799, 6.398, 5.815, WHITE, { r: 0.759, sh: SH_SOFT });
  box(s, 'ellipse', 3.699, 5.455, 1.068, 1.068, GREEN, { sh: SH_SOFT });
  box(s, 'ellipse', 0.429, 1.123, 3.392, 3.392, ORANGE, { sh: SH_SOFT });
  warehouse(s);
  tx(s, 'World\u2019s Leading Contract Logistic', 6.97, 1.11, 5.567, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, L1, 7.561, 3.036, 4.046, 0.686, { va: 'middle', c: GRAY });
  tx(s, 'Global Leaders', 7.561, 2.583, 3.95, 0.459, { va: 'middle', ls: 1.3, hd: true, c: ORANGE });
  box(s, 'roundRect', 6.97, 2.559, 0.508, 0.508, GREEN, { r: 0.254 });
  tx(s, L1, 7.561, 5.715, 4.344, 0.686, { va: 'middle', c: GRAY });
  tx(s, 'Brilliant Services', 7.561, 5.271, 3.95, 0.459, { va: 'middle', ls: 1.3, hd: true, c: ORANGE });
  box(s, 'roundRect', 6.97, 5.247, 0.508, 0.508, GREEN, { r: 0.254 });
  icon(s, 7.043, 5.326, 0.361, 0.361, CREAM);
  tx(s, L1, 7.561, 4.38, 4.344, 0.686, { va: 'middle', c: GRAY });
  tx(s, 'People Experience', 7.561, 3.927, 3.95, 0.459, { va: 'middle', ls: 1.3, hd: true, c: ORANGE });
  box(s, 'roundRect', 6.97, 3.903, 0.508, 0.508, GREEN, { r: 0.254 });
  box(s, 'roundRect', 5.902, -0.191, 0.597, 0.878, ORANGE, { r: 0.298 });
  box(s, 'chevron', 5.456, 5.977, 0.368, 0.45, ORANGE);
  box(s, 'chevron', 5.718, 5.977, 0.368, 0.45, CREAM);
  chrome(s, 24);
}

/** Slide 25 — proud to deliver excellence 95%. */
function slide25(pres) {
  const s = newSlide(pres);
  tx(s, 'Logistic', 8.982, 2.363, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  tx(s, 'Logistic', 5.565, 4.677, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  box(s, 'roundRect', 6.286, 5.182, 7.593, 1.171, GREEN, { r: 0.586, sh: SH_SOFT });
  tx(s, 'Proud to Deliver Excellence', 0.855, 1.658, 5.274, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, L3, 0.855, 4.788, 3.719, 1.122, { ls: 1.5, sz: 14, c: GRAY });
  box(s, 'ellipse', 6.286, 1.801, 2.901, 2.901, ORANGE, { sh: SH_SOFT });
  box(s, 'ellipse', 11.361, 4.429, 1.108, 1.108, ORANGE, { sh: SH_SOFT });
  airplane(s);
  box(s, 'roundRect', 6.378, -0.232, 0.597, 0.878, GREEN, { r: 0.298 });
  box(s, 'chevron', 11.331, 1.147, 0.368, 0.45, ORANGE);
  box(s, 'chevron', 11.594, 1.147, 0.368, 0.45, CREAM);
  tx(s, 'On-Time, Every Time', 0.855, 4.238, 3.719, 0.549, { ls: 1.5, sz: 20, hd: true, c: ORANGE });
  tx(s, '95%', 0.855, 2.911, 3.719, 1.445, { ls: 1.5, sz: 60, hd: true, c: GREEN });
  chrome(s, 25);
}

/** Slide 26 — key elements hub. */
function slide26(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 12.833, -0.232, 0.896, 8.003, ORANGE);
  tx(s, 'Logistic', 5.251, 6.67, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  box(s, 'ellipse', 7.206, 1.766, 4.148, 4.148, GREEN, { sh: SH_SOFT });
  box(s, 'ellipse', 6.038, 1.766, 1.82, 1.82, WHITE, { sh: SH_SOFT });
  box(s, 'ellipse', 8.369, 0.542, 1.82, 1.82, WHITE, { sh: SH_SOFT });
  tx(s, 'Logistics', 7.206, 3.375, 4.148, 0.774, { al: 'center', sz: 40, hd: true, c: WHITE });
  box(s, 'ellipse', 10.701, 1.766, 1.82, 1.82, WHITE, { sh: SH_SOFT });
  box(s, 'ellipse', 10.775, 4.106, 1.82, 1.82, WHITE, { sh: SH_SOFT });
  box(s, 'ellipse', 6.083, 4.093, 1.82, 1.82, WHITE, { sh: SH_SOFT });
  tx(s, 'Innovative Tech', 8.369, 1.306, 1.82, 0.707, { al: 'center', hd: true, c: ORANGE });
  box(s, 'ellipse', 8.324, 5.055, 1.82, 1.82, WHITE, { sh: SH_SOFT });
  tx(s, 'Key Elements of Our Logistics', 0.571, 2.425, 4.777, 1.313, { ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, L3, 0.542, 3.937, 4.653, 0.769, { ls: 1.5, sz: 14, c: GRAY });
  tx(s, 'Custom Solutions', 10.746, 2.522, 1.82, 0.707, { al: 'center', hd: true, c: ORANGE });
  tx(s, 'Experienced Team', 6.038, 2.515, 1.82, 0.707, { al: 'center', hd: true, c: ORANGE });
  tx(s, 'Flexible Options', 6.083, 4.864, 1.82, 0.707, { al: 'center', hd: true, c: ORANGE });
  tx(s, '24/7 Support', 8.408, 5.84, 1.652, 0.707, { al: 'center', hd: true, c: ORANGE });
  tx(s, 'Secure Shipping', 10.775, 4.864, 1.82, 0.707, { al: 'center', hd: true, c: ORANGE });
  box(s, 'roundRect', 9.026, 0.805, 0.508, 0.508, GREEN, { r: 0.254 });
  icon(s, 9.141, 0.927, 0.277, 0.263, CREAM);
  box(s, 'roundRect', 6.652, 2.007, 0.508, 0.508, GREEN, { r: 0.254 });
  box(s, 'roundRect', 6.74, 4.353, 0.508, 0.508, GREEN, { r: 0.254 });
  box(s, 'roundRect', 11.431, 4.359, 0.508, 0.508, GREEN, { r: 0.254 });
  box(s, 'roundRect', 11.388, 2.001, 0.508, 0.508, GREEN, { r: 0.254 });
  box(s, 'roundRect', 9.001, 5.292, 0.508, 0.508, GREEN, { r: 0.254 });
  icon(s, 11.51, 2.137, 0.265, 0.239, CREAM);
  icon(s, 6.753, 2.134, 0.307, 0.255, CREAM);
  box(s, 'rect', 11.524, 4.452, 0.322, 0.322, WHITE, { a: 100 });
  icon(s, 11.564, 4.465, 0.242, 0.282, CREAM);
  box(s, 'rect', 9.091, 5.383, 0.326, 0.326, WHITE, { a: 100 });
  icon(s, 9.132, 5.424, 0.245, 0.245, CREAM);
  icon(s, 6.859, 4.463, 0.269, 0.296, CREAM);
  box(s, 'roundRect', 6.378, -0.232, 0.597, 0.878, ORANGE, { r: 0.298 });
  tx(s, 'Logistic', 10.801, 0.508, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'chevron', 0.641, 5.144, 0.368, 0.45, ORANGE);
  box(s, 'chevron', 0.903, 5.144, 0.368, 0.45, CREAM);
  chrome(s, 26);
}

/** Slide 27 — full-service transportation. */
function slide27(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 1.579, 4.516, 0.909, 3.359, GREEN);
  box(s, 'rect', 5.107, 6.105, 0.909, 1.501, ORANGE);
  tx(s, 'Logistic', 5.322, 5.798, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'roundRect', 0.5, 1.064, 3.068, 3.849, ORANGE, { r: 0.511, sh: SH_SOFT });
  tx(s, 'End-to-End Solutions', 0.824, 1.992, 2.421, 1.318, { al: 'center', ls: 0.9, sb: 12, sz: 24, hd: true, c: WHITE });
  tx(s, 'Plan B', 1.333, 4.141, 1.403, 0.375, { al: 'center', va: 'middle', sz: 16, c: WHITE });
  tx(s, L1, 0.73, 2.977, 2.609, 1.026, { al: 'center', ls: 1.2, sz: 14, c: WHITE });
  box(s, 'roundRect', 4.027, 2.622, 3.068, 3.849, GREEN, { r: 0.511, sh: SH_SOFT });
  tx(s, 'Live Tracking', 4.563, 3.52, 1.997, 0.916, { al: 'center', ls: 0.9, sb: 12, sz: 24, hd: true, c: WHITE });
  tx(s, 'Plan C', 4.86, 5.699, 1.403, 0.375, { al: 'center', va: 'middle', sz: 16, c: WHITE });
  tx(s, L1, 4.355, 4.535, 2.413, 1.026, { al: 'center', ls: 1.2, sz: 14, c: WHITE });
  tx(s, 'Full-Service Transportation Logistics', 8.014, 1.916, 4.777, 2.121, { sz: 40, hd: true });
  tx(s, L2, 8.014, 4.279, 4.319, 1.122, { ls: 1.5, sz: 14, c: GRAY });
  box(s, 'roundRect', 8.014, -0.159, 0.597, 0.878, GREEN, { r: 0.298 });
  tx(s, 'Logistic', 10.801, 0.508, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  box(s, 'chevron', 4.113, 1.611, 0.368, 0.45, ORANGE);
  box(s, 'chevron', 4.376, 1.611, 0.368, 0.45, CREAM);
  icon(s, 5.336, 2.965, 0.452, 0.452, WHITE);
  icon(s, 1.721, 1.551, 0.625, 0.349, WHITE);
  chrome(s, 27);
}

/** Slide 28 — marketing mastery plan. */
function slide28(pres) {
  const s = newSlide(pres, WHITE);
  box(s, 'rect', 5.956, 5.264, 1.422, 3.395, GREEN, { sh: SH_DEEP });
  tx(s, 'Marketing Mastery Plan', 1.869, 0.834, 9.595, 0.707, { al: 'center', ls: 0.9, sb: 12, sz: 40, hd: true });
  tx(s, 'Logistic', 6.503, 3.769, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  tx(s, 'Logistic', 2.409, 3.722, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  petals(s);
  box(s, 'roundRect', 5.739, 3.418, 1.87, 1.87, WHITE, { r: 0.179, sh: SH_SOFT, rot: -45 });
  tx(s, '01', 6.404, 2.457, 0.498, 0.436, { al: 'center', nowrap: true, sz: 19.89, b: true, hd: true, c: WHITE });
  tx(s, '03', 7.982, 4.122, 0.553, 0.436, { al: 'center', nowrap: true, sz: 19.89, b: true, hd: true, c: WHITE });
  tx(s, '04', 6.384, 5.832, 0.584, 0.436, { al: 'center', nowrap: true, sz: 19.89, b: true, hd: true, c: WHITE });
  tx(s, '02', 4.798, 4.122, 0.554, 0.436, { al: 'center', nowrap: true, sz: 19.89, b: true, hd: true, c: WHITE });
  tx(s, L1 + ' ', 1.059, 2.788, 3.065, 0.769, { ls: 1.5, sz: 14, c: GRAY });
  tx(s, '1. Targeted Ads', 1.059, 2.453, 3.494, 0.404, { hd: true, c: GREEN });
  tx(s, L1 + ' ', 1.059, 5.27, 3.065, 0.769, { ls: 1.5, sz: 14, c: GRAY });
  tx(s, '2. Data Insights', 1.059, 4.935, 3.494, 0.404, { hd: true, c: ORANGE });
  tx(s, L1 + ' ', 9.162, 2.789, 3.065, 0.769, { ls: 1.5, sz: 14, c: GRAY });
  tx(s, '3. Customer Engagement', 9.162, 2.454, 3.671, 0.404, { hd: true, c: ORANGE });
  tx(s, L1 + ' ', 9.162, 5.27, 3.065, 0.769, { ls: 1.5, sz: 14, c: GRAY });
  tx(s, '4. Brand Consistency', 9.162, 4.935, 3.494, 0.404, { hd: true, c: GREEN });
  icon(s, 6.28, 3.853, 1, 1, ORANGE);
  box(s, 'roundRect', 12.205, -0.18, 0.597, 0.878, GREEN, { r: 0.298 });
  box(s, 'chevron', 2.094, 0.919, 0.368, 0.45, ORANGE);
  box(s, 'chevron', 2.356, 0.919, 0.368, 0.45, CREAM);
  chrome(s, 28);
}

/** Slide 29 — stay connected. */
function slide29(pres) {
  const s = newSlide(pres);
  box(s, 'rect', 6.401, 2.905, 7.155, 1.69, ORANGE);
  tx(s, 'Logistic', 4.617, 0.5, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_ORANGE, la: 80 });
  tx(s, 'Logistic', 4.419, 6.007, 4.167, 1.101, { ls: 0.9, sb: 12, sz: 66, hd: true, hollow: true, ln: WM_GREEN, la: 80 });
  box(s, 'roundRect', 6.208, 1.098, 6.126, 5.319, WHITE, { r: 0.694, sh: SH_TIGHT });
  tx(s, 'Stay Connected', 6.685, 1.406, 5.137, 0.774, { sz: 40, hd: true });
  box(s, 'chevron', 11.481, 0.852, 0.498, 0.498, ORANGE);
  box(s, 'chevron', 11.836, 0.852, 0.498, 0.498, CREAM);
  tx(s, 'www.yourwebsite.com', 7.466, 4.942, 4.493, 0.46, { ls: 1.5, sz: 16 });
  icon(s, 6.831, 5.035, 0.276, 0.274, GREEN);
  tx(s, '+123-456-7890', 7.466, 2.993, 4.493, 0.46, { ls: 1.5, sz: 16 });
  icon(s, 6.844, 3.099, 0.248, 0.248, GREEN);
  tx(s, 'Business Name', 7.466, 2.38, 4.493, 0.46, { ls: 1.5, sz: 16 });
  icon(s, 6.864, 2.474, 0.209, 0.26, GREEN);
  tx(s, 'hello@yourbusiness', 7.466, 3.647, 4.493, 0.46, { ls: 1.5, sz: 16 });
  icon(s, 6.835, 3.757, 0.267, 0.24, GREEN);
  tx(s, '@yournamebusiness', 7.466, 4.316, 4.493, 0.46, { ls: 1.5, sz: 16 });
  icon(s, 6.83, 4.422, 0.276, 0.248, GREEN);
  tx(s, '123 Anywhere St., Any City, ST 12345', 7.466, 5.593, 4.493, 0.46, { ls: 1.5, sz: 16 });
  icon(s, 6.739, 5.593, 0.459, 0.459, GREEN);
  box(s, 'roundRect', 6.208, -0.162, 0.597, 0.878, ORANGE, { r: 0.298 });
  box(s, 'line', 6.881, 2.84, 4.84, 0, null, { ln: CREAM });
  box(s, 'line', 6.881, 3.453, 4.84, 0, null, { ln: CREAM });
  box(s, 'line', 6.881, 4.124, 4.84, 0, null, { ln: CREAM });
  box(s, 'line', 6.881, 4.803, 4.84, 0, null, { ln: CREAM });
  box(s, 'line', 6.881, 5.424, 4.84, 0, null, { ln: CREAM });
  box(s, 'line', 6.881, 6.095, 4.84, 0, null, { ln: CREAM });
  chrome(s, 29);
}

/** Slide 30 — thank you. */
function slide30(pres) {
  const s = newSlide(pres, GREEN);
  box(s, 'rect', 0, 0, 13.333, 7.5, GREEN, { a: 40 });
  box(s, 'ellipse', 8.236, 3.137, 1.029, 1.029, ORANGE, { a: 20 });
  box(s, 'ellipse', 3.11, 2.164, 2.586, 2.586, ORANGE, { a: 20 });
  tx(s, 'For Your Attention', 4.583, 4.058, 4.167, 0.438, { al: 'center', sz: 20, c: CREAM });
  tx(s, 'Stay Connected!', 9.78, 6.896, 2.966, 0.404, { al: 'right', hd: true, c: CREAM });
  box(s, 'chevron', 7.872, 4.165, 0.268, 0.268, ORANGE);
  box(s, 'chevron', 8.064, 4.165, 0.268, 0.268, CREAM);
  tx(s, 'Thank You', 2.099, 2.948, 9.135, 1.111, { al: 'center', sz: 60, hd: true, c: CREAM });
  box(s, 'rect', 11.134, 0.003, 2.226, 0.845, ORANGE, { a: 20 });
  chrome(s, 30);
}

// ── build ───────────────────────────────────────────────────────────────────
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

function build() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'WIDE_13x7.5', width: 13.333, height: 7.5 });
  pres.layout = 'WIDE_13x7.5';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pres.title = 'Green Logistic';
  SLIDES.forEach(function (fn) { fn(pres); });
  return pres.writeFile({ fileName: path.join(__dirname, '061701b4-d25e-4178-8421-88de72e0ddec_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote ' + f); }).catch(function (e) { console.error(e); process.exit(1); });
