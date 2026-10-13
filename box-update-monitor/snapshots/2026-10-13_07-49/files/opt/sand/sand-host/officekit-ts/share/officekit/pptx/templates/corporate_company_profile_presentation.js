/**
 * NEWT - Company Presentation (40 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs. Photographs in the source deck are replaced by
 * flat gray placeholder shapes of the same position/size.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Palette / typography / boilerplate copy
 * ------------------------------------------------------------------ */
const C = {
  bg: '202020',      // slide background
  bar: '0D0D0D',     // top bar + dark cards
  panel: '262626',   // lighter dark panels
  white: 'FFFFFF',
  offWhite: 'F2F2F2',
  gray88: 'D8D8D8',
  gray75: 'BFBFBF',
  gray65: 'A5A5A5',
  gray50: '7F7F7F',
  green: '43E987',   // primary accent
  teal: '29D3CE',    // secondary accent
  gold: 'CDB320',    // tertiary accent
  photo: '989898',   // stand-in for photographs
  photoDark: '4E4E4E',
  mockup: 'B4B4B4'   // stand-in for device mock-ups
};

const F = {
  sans: 'Open Sans',
  light: 'Open Sans Light',
  semi: 'Open Sans SemiBold',
  xbold: 'Open Sans ExtraBold',
  body: 'Montserrat'
};

// The template's filler copy: one sentence chain, sliced at various lengths.
const L56 = 'Dev certa volle sopra anime animo qua. Tenta anima la no';
const L81 = L56 + ' aveva ch rombo le. Se da';
const L91 = L81 + ' distrutta';
const L111 = L91 + ' liberarli infantile';
const L130 = L111 + ' usignuoli. Ora afa';
const L142 = L130 + ' rimorso dai';
const L143 = L142 + '.';
const L181 = L142 + ' braccia sentito da distrutta liberarli';
const L202 = L181 + ' infantile usignuoli.';
const L196 = L142 + ' braccia sentito superbe chi. Indicibili ho esaltavano';
const L216 = L196 + ' raccontava un di fu';
const L227 = L196 + ' liberarli infantile usignuoli.';
const L252 = L216 + ' impregnato. Immemore un provarlo ho';
const L108 = 'Dev certa volle sopra anime animo qua. Tenta anima la aveva ch rombo le. Se da distrutta liberarli infantile';

const QUOTE_BOLD = 'Almost everything worthwhile carries with it some sort of risk, whether it\u2019s starting a new ';
const QUOTE_REST = 'business, whether it\u2019s leaving home, whether it\u2019s getting married, or whether it\u2019s flying into space';

const NOLINE = { type: 'none' };

// Shared chart styling: transparent plot on the dark cards, gray axes/gridlines.
const CHART_CATS = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
const CHART_DARK = {
  chartColors: [C.green, C.teal, C.gray65],
  showLegend: false,
  catAxisLabelColor: C.gray65, catAxisLabelFontFace: F.body, catAxisLabelFontSize: 10,
  catAxisLineColor: C.gray50,
  valAxisLabelColor: C.gray65, valAxisLabelFontFace: F.body, valAxisLabelFontSize: 10,
  valAxisLineShow: false,
  valGridLine: { color: C.gray50, size: 0.5, style: 'solid' },
  catGridLine: { style: 'none' },
  legendColor: C.gray65, legendFontFace: F.body, legendFontSize: 10,
  chartArea: { fill: { color: C.bar }, border: { pt: 0, color: C.bar } },
  plotArea: { fill: { color: C.bar } },
  border: { pt: 0, color: C.bar }
};

/* ------------------------------------------------------------------ *
 * Gradient helpers -- pptxgenjs has no gradient fill, so the deck's two
 * gradients are painted as a series of narrow solid strips / pie wedges.
 * ------------------------------------------------------------------ */
function mix(a, b, t) {
  const p = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map(i => Math.round(p(a, i) + (p(b, i) - p(a, i)) * t).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

// Colour of a multi-stop ramp at position t (stops = [[pos, hex], ...]).
function ramp(stops, t) {
  for (let i = 1; i < stops.length; i++) {
    if (t <= stops[i][0]) {
      const [p0, c0] = stops[i - 1], [p1, c1] = stops[i];
      return mix(c0, c1, p1 === p0 ? 0 : (t - p0) / (p1 - p0));
    }
  }
  return stops[stops.length - 1][1];
}

// "Accent" gradient used on every button/band: teal -> green, top-left to bottom-right.
const ACCENT = [[0, C.teal], [0.8, C.green], [1, C.green]];
// Ring gradient of the big cover circles: teal -> green -> gold.
const RING = [[0, C.teal], [0.2, C.teal], [0.56, C.green], [0.97, C.gold], [1, C.gold]];

// Diagonal accent gradient approximated with strips (green at start, teal at end).
function gradBand(s, x, y, w, h, opts) {
  const o = opts || {};
  const steps = o.steps || 24;
  const stops = o.stops || ACCENT;
  const vert = o.vertical;
  const span = vert ? h : w;
  for (let i = 0; i < steps; i++) {
    const t = 1 - i / (steps - 1);
    s.addShape('rect', {
      x: vert ? x : x + (span * i) / steps,
      y: vert ? y + (span * i) / steps : y,
      w: vert ? w : span / steps + 0.012,
      h: vert ? span / steps + 0.012 : h,
      fill: { color: ramp(stops, t) },
      line: NOLINE
    });
  }
}

// Same gradient inside a rounded pill: strips clipped by two end caps.
function gradPill(s, x, y, w, h, radius) {
  const r = Math.min(radius, h / 2, w / 2);
  gradBand(s, x + r, y, w - 2 * r, h, { steps: 18 });
  s.addShape('ellipse', { x: x, y: y + h / 2 - r, w: 2 * r, h: 2 * r, fill: { color: ramp(ACCENT, 1) }, line: NOLINE });
  s.addShape('ellipse', { x: x + w - 2 * r, y: y + h / 2 - r, w: 2 * r, h: 2 * r, fill: { color: ramp(ACCENT, 0) }, line: NOLINE });
}

// Circular teal->green->gold ring (cover slides) drawn from pie wedges.
// The ramp radiates from the bottom-right (teal) to the top-left (gold).
function gradRing(s, x, y, d) {
  const wedges = 36;
  for (let i = 0; i < wedges; i++) {
    const a0 = (i * 360) / wedges;
    const mid = a0 + 180 / wedges;
    const away = Math.abs(((mid - 135 + 540) % 360) - 180) / 180; // 0 at 135deg (bottom-right)
    s.addShape('pie', {
      x: x, y: y, w: d, h: d,
      angleRange: [a0, a0 + 360 / wedges + 0.6],
      fill: { color: ramp(RING, away) },
      line: NOLINE
    });
  }
}

/* ------------------------------------------------------------------ *
 * Shape / text helpers
 * ------------------------------------------------------------------ */
function rect(s, x, y, w, h, color, opts) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill: { color }, line: NOLINE }, opts || {}));
}

function hLine(s, x, y, w, color) {
  s.addShape('line', { x, y, w, h: 0, line: { color: color || C.gray65, width: 1 } });
}

function vLine(s, x, y, h, color) {
  s.addShape('line', { x, y, w: 0, h, line: { color: color || C.gray65, width: 1 } });
}

// Placeholder standing in for a photograph in the original deck.
function photo(s, x, y, w, h, opts) {
  const o = opts || {};
  s.addShape(o.shape || 'rect', {
    x, y, w, h,
    fill: { color: o.color || C.photo },
    line: o.line || NOLINE,
    rotate: o.rotate,
    rectRadius: o.rectRadius
  });
}

// 12pt Open Sans SemiBold / ExtraBold caption used all over the deck.
function label(s, text, x, y, w, opts) {
  const o = opts || {};
  s.addText(text, {
    x, y, w, h: o.h || 0.303,
    fontFace: o.font || F.xbold, fontSize: o.size || 12,
    color: o.color || C.offWhite, bold: o.bold, align: o.align || 'left',
    valign: 'top', lineSpacingMultiple: o.lnSpc
  });
}

// 11pt Montserrat paragraph, 1.5 line spacing (the deck's body style).
function body(s, text, x, y, w, h, opts) {
  const o = opts || {};
  s.addText(text, {
    x, y, w, h,
    fontFace: F.body, fontSize: o.size || 11,
    color: o.color || C.gray75, bold: o.bold, italic: o.italic,
    align: o.align || 'justify', valign: 'top',
    lineSpacingMultiple: 1.5
  });
}

// Two-tone 40pt headline: green/white, regular + light weights.
function heading(s, runs, x, y, w, opts) {
  const o = opts || {};
  s.addText(
    runs.map(r => ({
      text: r[0],
      options: { color: r[1] || C.white, fontFace: r[2] || F.sans, bold: r[3] !== false }
    })),
    { x, y, w, h: o.h || 1.447, fontSize: o.size || 40, align: o.align || 'left', valign: 'top' }
  );
}

// Big 48pt outline numeral / letter used beside list items.
function bigNum(s, text, x, y, opts) {
  const o = opts || {};
  s.addText(text, {
    x, y, w: o.w || 1.049, h: 0.909,
    fontFace: F.sans, fontSize: 48, bold: true,
    color: o.color || C.gray88, align: o.align || 'left', valign: 'top'
  });
}

// Filled (gradient) or outlined pill button with centred label.
function button(s, text, x, y, w, h, opts) {
  const o = opts || {};
  const r = (o.radiusFrac === undefined ? 0.41294 : o.radiusFrac) * Math.min(w, h);
  if (o.outline) {
    s.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: NOLINE, line: { color: o.outline, width: 1 } });
  } else if (o.solid) {
    s.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: { color: o.solid }, line: NOLINE });
  } else {
    gradPill(s, x, y, w, h, r);
  }
  s.addText(text, {
    x, y, w, h,
    fontFace: o.font || F.body, fontSize: o.size || 12, bold: o.bold !== false,
    color: o.color || C.white, align: 'center', valign: 'middle'
  });
}

// Flat (square) accent band with a centred caption -- "BEST TEAMS" style tags.
// The caption sits in its own top-anchored box, as in the source deck.
function tag(s, text, x, y, w, h, opts) {
  const o = opts || {};
  if (o.solid) rect(s, x, y, w, h, o.solid); else gradBand(s, x, y, w, h);
  const big = (o.size || 12) >= 20;
  s.addText(text, {
    x: o.tx === undefined ? x : o.tx,
    y: y + (big ? 0.027 : 0.099),
    w: o.tw === undefined ? w : o.tw,
    h: big ? 0.438 : 0.303,
    fontFace: F.xbold, fontSize: o.size || 12, color: o.color || C.offWhite,
    align: 'center', valign: 'top'
  });
}

// Green check bullet (circle + tick).
function check(s, x, y, d) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: C.green }, line: NOLINE });
  s.addText('\u2713', {
    x, y, w: d, h: d, fontFace: F.sans, fontSize: Math.round(d * 46), bold: true,
    color: C.white, align: 'center', valign: 'middle'
  });
}

// Decorative opening quotation mark -- one slab-serif comma, drawn twice.
// Outline traced from the source deck's freeform, normalised to a unit box.
const QUOTE_GLYPH = [
  ['M', 0.905, 0.268], ['C', 0.957, 0.268, 1.0, 0.248, 1.0, 0.224], ['L', 1.0, 0.044],
  ['C', 1.0, 0.02, 0.957, 0.0, 0.905, 0.0], ['C', 0.758, 0.0, 0.625, 0.015, 0.511, 0.043],
  ['C', 0.398, 0.072, 0.301, 0.113, 0.224, 0.165], ['C', 0.149, 0.216, 0.092, 0.277, 0.055, 0.347],
  ['C', 0.019, 0.415, 0.001, 0.493, 0.001, 0.577], ['L', 0.001, 0.956],
  ['C', 0.001, 0.98, 0.044, 1.0, 0.096, 1.0], ['L', 0.905, 1.0],
  ['C', 0.957, 1.0, 1.0, 0.98, 1.0, 0.956], ['L', 1.0, 0.577],
  ['C', 1.0, 0.552, 0.957, 0.533, 0.905, 0.533], ['L', 0.606, 0.533],
  ['C', 0.62, 0.357, 0.72, 0.268, 0.905, 0.268], ['Z']
];

function quoteMark(s, x, y, h, color) {
  const w = h * 0.463;         // one comma; the pair spans 1.075 * h
  [0, h * 0.612].forEach(dx => {
    s.addShape('custGeom', {
      x: x + dx, y, w, h,
      fill: { color: color || C.teal }, line: NOLINE,
      points: QUOTE_GLYPH.map(seg => {
        if (seg[0] === 'Z') return { close: true };
        if (seg[0] === 'C') {
          return { x: seg[5] * w, y: seg[6] * h, curve: { type: 'cubic', x1: seg[1] * w, y1: seg[2] * h, x2: seg[3] * w, y2: seg[4] * h } };
        }
        return { x: seg[1] * w, y: seg[2] * h, moveTo: seg[0] === 'M' };
      })
    });
  });
}

// Thin progress bar (track + filled portion).
function progress(s, x, y, w, h, frac) {
  rect(s, x, y, w, h, C.gray50);
  rect(s, x, y, w * frac, h, C.green);
}

// KPI card: dark rounded panel with "9,5K / Insert text here" columns.
function statCard(s, x, y, w, h, cols, perRow) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: 0.0617 * Math.min(w, h), fill: { color: C.bar }, line: NOLINE });
  const n = perRow || cols.length;
  const rows = Math.ceil(cols.length / n);
  cols.forEach((c, i) => {
    const cx = x + 0.051 + (i % n) * 1.764;
    const cy = y + 0.408 + Math.floor(i / n) * 1.076 + (rows > 1 ? 0 : 0);
    s.addText(c, { x: cx + 0.459, y: cy, w: 1.063, h: 0.37, fontFace: F.xbold, fontSize: 16, bold: true, color: C.green, align: 'center', valign: 'top' });
    label(s, 'Insert text here', cx, cy + 0.385, 1.982, { align: 'center' });
  });
}

/* ------------------------------------------------------------------ *
 * Page furniture
 * ------------------------------------------------------------------ */
function backdrop(s) {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  rect(s, 0, 0, 13.333, 0.706, C.bar);
}

function chrome(s, page) {
  label(s, '@newt.official.com', 0.58, 0.227, 1.75, { font: F.semi, color: C.gray88 });
  s.addText([{ text: 'Page' }, { text: '\t' }, { text: page }], {
    x: 11.564, y: 0.227, w: 1.516, h: 0.303,
    fontFace: F.semi, fontSize: 12, color: C.gray88, valign: 'top',
    tabStops: [{ position: 0.65 }]
  });
}

function signature(s) {
  rect(s, 0.463, 6.768, 0.117, 0.505, C.green);
  s.addText([
    { text: 'Yumnacreative', options: { breakLine: true } },
    { text: 'Present ' }
  ], { x: 0.58, y: 6.776, w: 1.415, h: 0.505, fontFace: F.semi, fontSize: 12, color: C.gray88, valign: 'top' });
}

// Rotated "2020 DESIGN" caption plus its vertical rule at the right edge.
function sideLabel(s, textY, lineY) {
  s.addText('2020 DESIGN', {
    x: 12.266, y: textY, w: 1.247, h: 0.303, rotate: 90,
    fontFace: F.semi, fontSize: 12, color: C.gray88, valign: 'top'
  });
  vLine(s, 12.889, lineY, 1.678);
}

/* ------------------------------------------------------------------ *
 * Slide builders
 * ------------------------------------------------------------------ */
const slides = [];

// 01 - Cover
slides.push(s => {
  backdrop(s);
  s.addShape('parallelogram', { x: 6.796, y: 0.706, w: 5.81, h: 6.794, fill: { color: C.panel }, line: NOLINE });
  s.addShape('parallelogram', { x: 0.58, y: 0.706, w: 5.81, h: 6.794, fill: { color: C.panel }, line: NOLINE, flipH: true });
  gradRing(s, 3.909, 1.207, 5.515);
  s.addShape('ellipse', { x: 4.048, y: 1.345, w: 5.238, h: 5.238, fill: { color: C.bar }, line: NOLINE });
  s.addShape('ellipse', { x: 4.177, y: 1.475, w: 4.98, h: 4.98, fill: NOLINE, line: { color: C.white, width: 1 } });
  s.addShape('ellipse', { x: 2.795, y: 1.624, w: 0.774, h: 0.774, fill: { color: ramp(ACCENT, 0.45) }, line: NOLINE });
  s.addShape('ellipse', { x: 9.884, y: 5.512, w: 0.261, h: 0.261, fill: { color: ramp(ACCENT, 0.45) }, line: NOLINE });
  chrome(s, '01');
  s.addText('NEWT', { x: 3.764, y: 2.908, w: 5.805, h: 1.447, fontSize: 80, bold: true, fontFace: F.sans, color: C.white, align: 'center', valign: 'top' });
  gradPill(s, 4.772, 4.355, 3.79, 0.451, 0.41294 * 0.451);
  label(s, 'COMPANY PRESENTATION', 4.55, 4.397, 4.246, { align: 'center', bold: true, color: C.white, h: 0.37, lnSpc: 1.5 });
  signature(s);
  sideLabel(s, 4.486, 5.595);
});

// 02 - About (circular photo)
slides.push(s => {
  backdrop(s);
  s.addShape('parallelogram', { x: 6.1, y: 0.706, w: 5.81, h: 6.794, fill: { color: C.panel }, line: NOLINE, flipH: true });
  gradRing(s, 0.794, 1.345, 5.515);
  photo(s, 1.071, 1.345, 5.238, 5.238, { shape: 'ellipse' });
  chrome(s, '02');
  heading(s, [['ABOUT NEWT ', C.green], ['COMPANY', C.white, F.light]], 7.097, 1.858, 4.871);
  body(s, L252, 7.097, 3.502, 5.365, 1.182);
  label(s, 'INSERT YOUR TEXT', 7.097, 5.642, 1.907);
  body(s, L130, 7.097, 5.953, 5.365, 0.625, { color: C.gray50 });
  hLine(s, 6.031, 5.794, 0.993);
});

// 03 - Quote with side photo
slides.push(s => {
  backdrop(s);
  quoteMark(s, 1.253, 1.989, 0.833);
  s.addText([
    { text: QUOTE_BOLD, options: { bold: true } },
    { text: QUOTE_REST }
  ], { x: 1.239, y: 2.872, w: 6.022, h: 2.524, fontFace: F.sans, fontSize: 24, color: C.white, valign: 'top' });
  tag(s, 'NEWT COMPANY', 8.457, 2.071, 2.814, 0.5);
  chrome(s, '03');
  signature(s);
  sideLabel(s, 4.486, 5.595);
  photo(s, 8.457, 2.854, 2.814, 2.499);
});

// 04 - About with stacked photos
slides.push(s => {
  backdrop(s);
  gradBand(s, 0.5, 5.529, 2.814, 1.229);
  chrome(s, '04');
  photo(s, 0.5, 1.251, 2.814, 4.064);
  photo(s, 3.557, 2.829, 3.543, 3.929);
  heading(s, [['ABOUT NEWT ', C.green], ['COMPANY', C.white, F.light]], 7.334, 2.011, 4.871);
  body(s, L130, 7.468, 3.597, 5.365, 0.625);
  label(s, 'INSERT YOUR TEXT', 7.468, 4.543, 1.907);
  hLine(s, 9.389, 4.695, 1.524);
  body(s, L196, 7.468, 4.863, 5.365, 0.903, { color: C.gray50 });
  label(s, 'NEWT COMPANY', 0.953, 5.639, 1.907, { align: 'center' });
  body(s, L56, 0.625, 5.942, 2.563, 0.625, { color: C.offWhite, align: 'center' });
});

// 05 - Build your Impression
slides.push(s => {
  backdrop(s);
  rect(s, 6.1, 0.706, 5.81, 6.794, C.panel);
  chrome(s, '05');
  [[1.151, C.gray50], [2.299, C.gray65], [3.447, C.gray75]].forEach(([y, col]) =>
    label(s, 'INSERT TEXT HERE', 0.37, y, 1.96, { align: 'center', color: col })
  );
  photo(s, 3.443, 1.036, 6.886, 2.714);
  heading(s, [['Build your Impression ', C.green], ['with our company', C.white, F.light]], 2.954, 4.011, 7.426, { align: 'center' });
  body(s, L252, 2.975, 5.622, 7.383, 0.903, { align: 'center' });
  button(s, 'READ MORE', 10.358, 6.734, 2.411, 0.475);
});

// 06 - Counting Impression
slides.push(s => {
  backdrop(s);
  chrome(s, '06');
  photo(s, 7.243, 0.706, 4.829, 4.965);
  label(s, 'COMPANY PROFILE', 0.88, 1.587, 1.907);
  heading(s, [['Counting ', C.green], ['Impression', C.white, F.light]], 0.829, 1.841, 5.838, { h: 0.774 });
  body(s, L252, 0.829, 2.819, 5.838, 1.182);
  statCard(s, 0.829, 4.528, 5.627, 1.504, ['9,5K', '10M', '240']);
  button(s, 'CONTINUE', 10.358, 6.724, 2.411, 0.496, { outline: C.green });
});

// 07 - History
slides.push(s => {
  backdrop(s);
  gradRing(s, 0.498, 1.156, 4.381);
  s.addShape('ellipse', { x: 0.718, y: 1.384, w: 4.16, h: 4.16, fill: { color: C.bar }, line: NOLINE });
  chrome(s, '07');
  heading(s, [['HISTORY NEWT ', C.green], ['COMPANY', C.white, F.light]], 1.255, 2.467, 4.202, { h: 2.121 });
  photo(s, 5.457, 1.184, 3.271, 2.566);
  photo(s, 9.143, 1.184, 3.271, 2.566);
  hLine(s, 4.447, 4.169, 1.524);
  label(s, 'INSERT YOUR TEXT', 5.457, 4.285, 1.907);
  body(s, L252, 5.457, 4.588, 6.957, 0.903);
  [[1.255, 1.905, '1'], [7.015, 7.665, '2']].forEach(([nx, tx, n]) => {
    bigNum(s, n, nx, 5.953, { w: 0.514, color: C.gray65 });
    body(s, L111, tx, 6.015, 4.75, 0.625, { color: C.gray50 });
  });
});

// 08 - Mission
slides.push(s => {
  backdrop(s);
  rect(s, 0, 1.843, 8.814, 3.857, C.panel);
  chrome(s, '08');
  photo(s, 8.814, 0.706, 4.519, 6.794);
  heading(s, [['Mission ', C.green], ['of Newt', C.white, F.light]], 0.997, 1.352, 5.838, { h: 0.774 });
  label(s, 'MISSION - COMPANY PROFILE', 1.048, 2.068, 3.018);
  body(s, L216, 1.01, 2.486, 6.158, 0.903);
  [3.892, 5.028].forEach(y => {
    check(s, 1.048, y, 0.357);
    label(s, 'INSERT YOUR TEXT', 1.478, y, 1.907);
    body(s, L142, 1.478, y + 0.32, 5.69, 0.625, { color: C.gray50 });
  });
  hLine(s, 5.367, 6.148, 1.524);
});

// 09 - Vision
slides.push(s => {
  backdrop(s);
  rect(s, 6.833, 1.514, 5.442, 1.312, C.bar);
  gradBand(s, 12.014, 2.499, 0.435, 0.408, { steps: 6 });
  chrome(s, '09');
  photo(s, 0.537, 1.514, 2.863, 5.178);
  photo(s, 3.637, 1.514, 2.863, 5.178);
  heading(s, [['Vision ', C.green], ['of Newt', C.white, F.light]], 7.101, 1.584, 4.957, { h: 0.774 });
  label(s, 'VISION - COMPANY PROFILE', 7.153, 2.299, 3.018);
  body(s, L202, 6.877, 3.073, 5.529, 0.903);
  [['01', 4.338], ['02', 5.536]].forEach(([n, y]) => {
    bigNum(s, n, 6.781, y);
    label(s, 'INSERT YOUR TEXT', 7.83, y + 0.109, 1.907);
    body(s, L111, 7.83, y + 0.429, 4.576, 0.625, { color: C.gray50 });
  });
});

// 10 - Team
slides.push(s => {
  backdrop(s);
  rect(s, 0.352, 4.185, 2.677, 3.0, C.panel);
  rect(s, 3.509, 1.021, 2.677, 3.0, C.bar);
  chrome(s, '10');
  photo(s, 0.352, 1.021, 2.677, 3.0);
  photo(s, 3.509, 4.185, 2.677, 3.0);
  // person card: left-aligned (top) and right-aligned (bottom)
  label(s, 'Job title here', 3.639, 1.54, 2.094, { font: F.light, color: C.white });
  label(s, 'Laura Taylor', 3.639, 1.912, 2.094, { size: 14, bold: true, color: C.green, h: 0.337 });
  body(s, L108, 3.639, 2.248, 2.546, 1.18, { color: C.gray50, align: 'left' });
  label(s, 'Job title here', 0.848, 4.771, 2.094, { font: F.light, color: C.white, align: 'right' });
  label(s, 'Wayne Russell', 0.848, 5.144, 2.094, { size: 14, bold: true, color: C.green, h: 0.337, align: 'right' });
  body(s, L111, 0.352, 5.48, 2.59, 1.18, { color: C.gray50, align: 'right' });
  heading(s, [['TEAM NEWT ', C.green], ['COMPANY', C.white, F.light]], 7.104, 2.651, 4.229);
  body(s, L196, 7.104, 4.237, 5.153, 0.903);
  hLine(s, 10.572, 5.482, 1.524);
  button(s, 'DETAIL HERE', 7.104, 6.464, 2.411, 0.394, { outline: C.green });
});

// 11 - Manager profile with skill bars
slides.push(s => {
  backdrop(s);
  rect(s, 0, 2.36, 13.333, 5.14, C.panel);
  chrome(s, '11');
  photo(s, 8.918, 1.325, 3.254, 4.892);
  label(s, 'MANAGER OF NEWT COMPANY', 1.048, 1.209, 3.018);
  heading(s, [['Roger ', C.green], ['Sanchez', C.white, F.sans, false]], 0.997, 1.433, 5.838, { h: 0.774 });
  quoteMark(s, 1.18, 2.462, 0.318, C.green);
  body(s, L143, 1.048, 2.836, 6.158, 0.625, { bold: true });
  body(s, L227, 1.048, 3.726, 6.158, 0.903, { color: C.gray50 });
  [['45%', 4.925, 0.357], ['85%', 5.73, 0.723]].forEach(([pct, y, frac]) => {
    label(s, 'WRITE YOUR TEXT', 1.027, y + 0.017, 1.603, { size: 11, bold: true, color: C.gray75, h: 0.286 });
    label(s, pct, 6.196, y, 0.558, { bold: true, color: C.gray75 });
    progress(s, 1.131, y + 0.375, 5.805, 0.107, frac);
  });
  sideLabel(s, 4.486, 5.595);
});

// 12 - Expert team grid
slides.push(s => {
  backdrop(s);
  s.addShape('parallelogram', { x: 3.557, y: 0.706, w: 5.81, h: 6.794, fill: { color: C.panel }, line: NOLINE, flipH: true });
  chrome(s, '12');
  const team = [
    ['EUGENE BROWN', 0.823, 1.364, 0.885, 4.104],
    ['NICHOLAS BENNETT', 3.795, 1.364, 3.856, 4.104],
    ['BRENDA JONES', 7.005, 3.75, 7.066, 6.361],
    ['KATHLEEN CLARK', 9.976, 3.75, 10.038, 6.361]
  ];
  team.forEach(([name, px, py, bx, by]) => {
    photo(s, px, py, 2.534, 3.093);
    gradPill(s, bx, by, 2.411, 0.911, 0.19028 * 0.911);
    label(s, name, bx + 0.031, by + 0.113, 2.38, { size: 14, bold: true, color: C.white, align: 'center', h: 0.337 });
    label(s, 'Job description', bx + 0.174, by + 0.45, 2.094, { font: F.light, color: C.white, align: 'center' });
  });
  heading(s, [['EXPERT TEAM ', C.green], ['NEWT', C.white, F.sans, false], [' ', C.green], ['COMPANY', C.white, F.light]], 7.198, 1.54, 5.106);
  hLine(s, 10.12, 3.169, 1.524);
  label(s, 'INSERT YOUR TEXT', 1.059, 5.523, 1.907);
  body(s, L111, 1.059, 5.843, 4.576, 0.625, { color: C.gray50 });
});

// 13 - Best teams (4 columns)
slides.push(s => {
  backdrop(s);
  rect(s, 0, 1.94, 13.333, 3.7, C.panel);
  chrome(s, '13');
  ['Donna Allen', 'Harold Phillips', 'Jason Young', 'Sandra Hughes'].forEach((name, i) => {
    const x = [1.009, 3.866, 6.667, 9.58][i];
    label(s, 'Job title here', x, 2.141, 2.094, { font: F.light, color: C.white });
    label(s, name, x, 2.513, 2.094, { size: 14, bold: true, color: C.green, h: 0.337 });
    body(s, L108, x, 2.879, 2.546, 1.18, { color: C.gray50, align: 'left' });
    photo(s, [1.009, 3.866, 6.723, 9.58][i], 4.272, 2.448, 2.736);
  });
  tag(s, 'BEST TEAMS', 4.24, 0.8, 2.814, 0.5, { size: 20, tx: 4.693, tw: 1.907 });
  tag(s, 'HELP YOUR BUSINESS', 5.431, 1.3, 2.814, 0.5, { solid: C.bar, tx: 5.733, tw: 2.212 });
});

// 14 - Two profiles
slides.push(s => {
  backdrop(s);
  rect(s, 6.667, 4.122, 5.514, 2.547, C.bar);
  chrome(s, '14');
  photo(s, 6.667, 1.556, 5.324, 2.547);
  photo(s, 1.342, 4.103, 5.324, 2.547);
  label(s, 'INSERT YOUR TEXT', 1.3, 1.695, 1.907);
  body(s, L181, 1.3, 1.955, 5.007, 0.903);
  body(s, L111, 1.3, 2.978, 4.95, 0.625, { color: C.gray50 });
  tag(s, 'BEVERLY MURPHY', 0.58, 3.853, 2.814, 0.5, { size: 20 });
  tag(s, 'BRANDON TORRES', 9.99, 3.853, 2.814, 0.5, { size: 20 });
  label(s, 'INSERT YOUR TEXT', 7.003, 4.493, 1.907);
  body(s, L181, 7.003, 4.752, 5.007, 0.903);
  body(s, L111, 7.003, 5.776, 4.95, 0.625, { color: C.gray50 });
});

// 15 - Expert service (6 check items)
slides.push(s => {
  backdrop(s);
  rect(s, 0, 2.167, 2.625, 5.333, C.bar);
  chrome(s, '15');
  hLine(s, 3.446, 2.679, 1.524);
  heading(s, [['EXPERT SERVICE ', C.green], ['NEWT', C.white, F.sans, false], [' ', C.green], ['COMPANY', C.white, F.light]], 0.606, 2.861, 5.106);
  tag(s, 'EXPERT SERVICE', 0.755, 4.674, 2.814, 0.5, { size: 20, tx: 0.879, tw: 2.567 });
  [1.456, 3.262, 5.067].forEach(y => {
    [5.776, 9.342].forEach(x => {
      check(s, x, y, 0.357);
      label(s, 'INSERT YOUR TEXT', x + 0.429, y, 1.907);
      body(s, L91, x + 0.429, y + 0.32, 2.809, 0.903, { color: C.gray50 });
    });
  });
  signature(s);
});

// 16 - Best service (numbered cards)
slides.push(s => {
  backdrop(s);
  gradBand(s, 1.278, 1.222, 1.528, 2.653);
  rect(s, 4.89, 2.59, 1.528, 2.653, C.bar);
  rect(s, 1.296, 4.219, 1.528, 2.653, C.bar);
  chrome(s, '16');
  const cards = [
    ['01', 2.087, 1.341, 0.918, C.white, C.offWhite],
    ['02', 5.681, 2.709, 4.512, C.gray88, C.gray50],
    ['03', 2.087, 4.337, 0.918, C.gray88, C.gray50]
  ];
  cards.forEach(([n, nx, ny, bx, numCol, txtCol]) => {
    bigNum(s, n, nx, ny, { align: 'center', color: numCol });
    label(s, 'INSERT YOUR TEXT', bx + 0.74, ny + 0.926, 1.907, { align: 'center', color: numCol === C.white ? C.white : C.offWhite });
    body(s, L111, bx, ny + 1.245, 3.387, 0.903, { color: txtCol, align: 'center' });
  });
  label(s, 'SERVICE OF NEWT COMPANY', 8.553, 2.877, 3.018);
  heading(s, [['Best ', C.green], ['Service', C.white, F.sans, false]], 8.502, 3.101, 4.29, { h: 0.774 });
  body(s, L142, 8.502, 3.889, 4.04, 0.903);
  button(s, 'SERVICE PAGE', 8.502, 5.02, 2.411, 0.394, { outline: C.green });
});

// 17 - Timeline
slides.push(s => {
  backdrop(s);
  gradBand(s, 0, 6.714, 13.333, 0.159);
  chrome(s, '17');
  heading(s, [['EXPERT SERVICE ', C.green], ['NEWT', C.white, F.sans, false], [' ', C.green], ['COMPANY', C.white, F.light]], 4.114, 1.014, 5.106, { align: 'center' });
  button(s, 'SERVICE PAGE', 5.461, 2.645, 2.411, 0.394, { outline: C.green });
  [['01', 1.853, 1.094, 2.191, 5.872], ['02', 4.742, 3.983, 5.128, 6.011],
   ['03', 7.631, 6.872, 8.017, 5.872], ['04', 10.52, 9.761, 10.951, 6.011]].forEach(([n, nx, bx, dotX, dotY]) => {
    bigNum(s, n, nx, 3.875, { align: 'center' });
    body(s, L81, bx, 4.79, 2.568, 0.903, { color: C.gray50, align: 'center' });
    vLine(s, dotX + 0.139, 6.149, 0.559);
    s.addShape('ellipse', { x: dotX, y: dotY, w: 0.278, h: 0.278, fill: { color: ramp(ACCENT, 0.5) }, line: NOLINE });
  });
});

// 18 - Two services split by an accent rule
slides.push(s => {
  backdrop(s);
  gradBand(s, 0, 4.024, 13.333, 0.159);
  chrome(s, '18');
  photo(s, 9.257, 1.475, 2.734, 2.547);
  photo(s, 1.342, 4.183, 2.734, 2.547);
  label(s, 'INSERT YOUR TEXT', 1.999, 1.475, 1.907);
  body(s, L252, 1.999, 1.778, 6.957, 0.903);
  tag(s, 'SERVICE ONE', 6.443, 3.5, 2.814, 0.5, { solid: C.bar, tx: 6.744, tw: 2.212 });
  gradBand(s, 6.151, 3.124, 0.749, 0.635, { steps: 8 });
  bigNum(s, '01', 5.627, 2.761, { align: 'center', color: C.white });
  tag(s, 'SERVICE TWO', 4.076, 4.219, 2.814, 0.5, { solid: C.bar, tx: 4.377, tw: 2.212 });
  gradBand(s, 6.443, 4.478, 0.749, 0.635, { steps: 8 });
  bigNum(s, '02', 6.589, 4.561, { align: 'center', color: C.white });
  label(s, 'INSERT YOUR TEXT', 4.377, 5.53, 1.907);
  body(s, L252, 4.377, 5.833, 6.957, 0.903);
});

// 19 - Best service with vertical accent rail
slides.push(s => {
  backdrop(s);
  gradBand(s, 7.759 + (1.001 - 0.172) / 2, 0.706, 0.172, 6.794, { vertical: true });
  rect(s, 0, 5.343, 2.129, 2.157, C.bar);
  chrome(s, '19');
  photo(s, 0.714, 2.514, 3.686, 3.995);
  label(s, 'SERVICE OF NEWT COMPANY', 0.87, 1.291, 3.018);
  heading(s, [['Best ', C.green], ['Service', C.white, F.sans, false]], 0.819, 1.515, 4.29, { h: 0.774 });
  [1.329, 3.468, 5.574].forEach((y, i) => {
    rect(s, 8.993, y - 0.359, 3.771, 1.931, C.bar);
    check(s, 9.248, y, 0.357);
    label(s, 'INSERT YOUR TEXT', 9.677, y, 1.907);
    body(s, L91, 9.677, y + 0.32, 2.809, 0.903, { color: C.gray50 });
    // dot + tick marks on the rail
    vLine(s, 8.203, [1.782, 3.935, 6.094][i], 0.559);
    s.addShape('ellipse', { x: [8.482, 8.343, 8.482][i], y: [1.92, 4.076, 6.232][i], w: 0.278, h: 0.278, fill: { color: ramp(ACCENT, 0.5) }, line: NOLINE });
  });
  tag(s, 'BEST SERVICE', 2.642, 5.578, 2.814, 0.5, { size: 20 });
  tag(s, 'HELP YOUR IMPRESSION', 3.833, 6.078, 2.814, 0.5, { solid: C.bar });
});

// 20 - Breaktime
slides.push(s => {
  backdrop(s);
  rect(s, 0.589, 1.122, 5.168, 5.168, C.bar);
  chrome(s, '20');
  photo(s, 1.378, 3.06, 6.879, 3.74);
  s.addText('BREAKTIME', { x: 1.455, y: 1.412, w: 7.236, h: 1.447, fontFace: F.sans, fontSize: 80, bold: true, color: C.white, valign: 'top' });
  label(s, 'WE\u2019LL BE BACK ON 10 MINUTES', 8.556, 3.525, 2.967);
  body(s, L142, 8.556, 3.828, 3.882, 0.903);
  gradBand(s, 7.118, 5.033, 4.406, 1.422);
  s.addText('Enjoy your lunch & breaktime before we continue our presentation', {
    x: 7.308, y: 5.162, w: 4.017, h: 1.111, fontFace: F.xbold, fontSize: 20, color: C.offWhite, valign: 'top'
  });
  signature(s);
  sideLabel(s, 3.09, 0.842);
});

// 21 - Best portfolio
slides.push(s => {
  backdrop(s);
  gradBand(s, 11.032, 0.706, 2.302, 5.167, { vertical: true });
  chrome(s, '21');
  heading(s, [['BEST PORTFOLIO ', C.green], ['NEWT', C.white, F.sans, false], [' ', C.green], ['COMPANY', C.white, F.light]], 0.832, 1.243, 5.106);
  hLine(s, 0.989, 2.906, 1.524);
  body(s, L91, 0.925, 3.08, 4.313, 0.625, { color: C.gray50 });
  tag(s, 'COLLABORATION WORKS', 4.328, 4.145, 4.313, 0.5, { size: 20 });
  photo(s, 8.952, 1.467, 3.686, 3.995);
  [0.63, 3.284, 5.938].forEach((x, i) => photo(s, x, 4.81, 2.654, 2.382, { color: i === 1 ? C.photoDark : C.photo }));
});

// 22 - Portfolio columns
slides.push(s => {
  backdrop(s);
  chrome(s, '22');
  [[0, 1.514], [2.857, 1.514], [5.714, 1.514]].forEach(([x, y]) => photo(s, x, y, 2.614, 5.178));
  // caption bands: two on top of their column, one at the bottom
  [[0, 1.514], [5.714, 1.514], [2.857, 5.463]].forEach(([x, y]) => {
    gradBand(s, x, y, 2.614, 1.229);
    label(s, 'NEWT PROFILE', x + 0.379, y + 0.161, 1.907, { align: 'center' });
    body(s, L56, x + 0.051, y + 0.464, 2.563, 0.625, { color: C.offWhite, align: 'center' });
  });
  label(s, 'PORTFOLIO OF NEWT COMPANY', 8.623, 2.892, 3.018);
  heading(s, [['Best ', C.green], ['Portfolio', C.white]], 8.571, 3.115, 4.29, { h: 0.774 });
  body(s, L142, 8.623, 3.991, 3.882, 0.903);
  body(s, L91, 8.623, 4.995, 3.882, 0.625);
});

// 23 - Portfolio page
slides.push(s => {
  backdrop(s);
  chrome(s, '23');
  button(s, 'PORTFOLIO PAGE', 0.45, 1.439, 2.411, 0.394, { outline: C.green });
  heading(s, [['Best ', C.green], ['Portfolio', C.white]], 0.45, 1.978, 4.29, { h: 0.774 });
  label(s, 'PORTFOLIO OF NEWT COMPANY', 0.502, 2.752, 3.018);
  [3.344, 4.615].forEach(y => {
    label(s, 'INSERT YOUR TEXT', 0.502, y, 1.907);
    body(s, L181, 0.502, y + 0.26, 5.007, 0.903);
  });
  photo(s, 6.667, 1.167, 6.216, 3.583);
  [5.914, 9.309].forEach(x => photo(s, x, 4.058, 2.833, 2.595, { line: { color: C.bar, width: 6 } }));
});

// 24 - Portfolio grid (4 framed photos)
slides.push(s => {
  backdrop(s);
  gradBand(s, 7.181, 1.368, 3.972, 1.194);
  chrome(s, '24');
  [[0.553, 1.368], [3.691, 1.368], [0.553, 4.243], [3.691, 4.243]].forEach(([x, y]) =>
    photo(s, x, y, 2.833, 2.595, { line: { color: C.bar, width: 6 } }));
  heading(s, [['BEST PORTFOLIO ', C.white], ['NEWT', C.white, F.sans, false], [' ', C.white], ['COMPANY', C.white, F.light]], 7.471, 1.784, 5.106);
  quoteMark(s, 7.307, 3.251, 0.318, C.green);
  body(s, L111, 7.175, 3.626, 5.106, 0.625, { bold: true });
  [['1', 4.591], ['2', 5.856]].forEach(([n, y]) => {
    bigNum(s, n, 7.067, y + 0.178);
    label(s, 'INSERT YOUR TEXT', 7.908, y, 1.907);
    body(s, L91, 7.908, y + 0.32, 4.374, 0.625, { color: C.gray50 });
  });
});

// 25 - Portfolio banner over photo strips
slides.push(s => {
  backdrop(s);
  chrome(s, '25');
  [[0, 2.614], [2.614, 2.614], [5.229, 2.614], [7.843, 2.614], [10.457, 2.876]].forEach(([x, w], i) =>
    photo(s, x, 0.706, w, 6.794, { color: i % 2 ? C.photoDark : C.photo }));
  rect(s, 4.188, 4.782, 5.934, 1.454, C.bar);
  gradBand(s, 2.95, 3.569, 4.938, 1.213);
  heading(s, [['BEST PORTFOLIO ', C.white], ['NEWT', C.white, F.sans, false], [' ', C.white], ['COMPANY', C.white, F.light]], 3.216, 3.159, 5.106);
  quoteMark(s, 4.665, 5.022, 0.318, C.green);
  body(s, L111, 4.533, 5.397, 5.106, 0.625, { bold: true });
  gradBand(s, 9.841, 4.535, 0.502, 0.502, { steps: 7 });
  signature(s);
  sideLabel(s, 4.828, 5.603);
});

// 26 - Portfolio mosaic
slides.push(s => {
  backdrop(s);
  chrome(s, '26');
  [6.011, 8.289, 10.566].forEach(x => photo(s, x, 1.82, 2.003, 1.834));
  photo(s, 3.747, 3.84, 2.003, 1.834);
  photo(s, 6.011, 3.848, 6.559, 1.834);
  label(s, 'PORTFOLIO OF NEWT COMPANY', 0.631, 1.278, 3.018);
  heading(s, [['Best ', C.green], ['Portfolio', C.white]], 0.58, 1.581, 4.29, { h: 0.774 });
  body(s, L181, 0.58, 2.528, 5.007, 0.903);
  label(s, 'INSERT YOUR TEXT', 0.58, 4.222, 1.907);
  body(s, L91, 0.58, 4.541, 2.906, 0.903, { color: C.gray50 });
  label(s, 'INSERT YOUR TEXT', 0.631, 5.812, 1.907);
  body(s, L252, 0.631, 6.115, 6.957, 0.903, { color: C.gray50 });
  button(s, 'READ MORE', 8.014, 5.444, 2.411, 0.475);
});

// 27 - Desktop mockup
slides.push(s => {
  backdrop(s);
  gradBand(s, 7.143, 1.286, 1.671, 2.8);
  chrome(s, '27');
  // monitor stand-in: bezel, screen, foot
  s.addShape('roundRect', { x: 1.006, y: 1.186, w: 5.129, h: 3.29, rectRadius: 0.08, fill: { color: '1F1F1F' }, line: NOLINE });
  photo(s, 1.186, 1.36, 4.777, 2.777);
  rect(s, 1.006, 4.31, 5.129, 0.5, C.mockup);
  rect(s, 3.1, 4.81, 0.94, 0.9, 'A8A8A8');
  s.addShape('ellipse', { x: 2.55, y: 5.55, w: 2.04, h: 0.44, fill: { color: 'BFBFBF' }, line: NOLINE });
  heading(s, [['BEST DESKTOP ', C.white], ['NEWT', C.white, F.sans, false], [' ', C.white], ['COMPANY', C.white, F.light]], 7.471, 1.988, 5.106);
  label(s, 'INSERT YOUR TEXT', 7.023, 4.373, 1.907);
  body(s, L181, 7.023, 4.633, 5.007, 0.903);
  signature(s);
});

// 28 - Mobile mockup
slides.push(s => {
  backdrop(s);
  chrome(s, '28');
  [7.731, 10.146].forEach(x => {
    s.addShape('roundRect', { x, y: 1.794, w: 2.015, h: 4.02, rectRadius: 0.28, fill: { color: '3A3A3A' }, line: NOLINE });
    photo(s, x + 0.112, 1.906, 1.791, 3.796, { shape: 'roundRect', rectRadius: 0.2 });
  });
  gradBand(s, 1.124, 1.929, 4.335, 1.213);
  heading(s, [['BEST MOBILE ', C.white], ['NEWT', C.white, F.sans, false], [' ', C.white], ['COMPANY', C.white, F.light]], 1.794, 1.518, 5.106);
  quoteMark(s, 1.208, 3.552, 0.318, C.green);
  body(s, L143, 1.076, 3.926, 5.106, 0.903, { bold: true });
  button(s, 'DOWNLOAD', 1.076, 5.315, 2.411, 0.475, { solid: C.bar });
  button(s, 'INSTALL', 3.763, 5.315, 2.411, 0.475, { solid: C.panel });
});

// 29 - Watch mockup
slides.push(s => {
  backdrop(s);
  chrome(s, '29');
  // watch stand-in: strap, body, screen
  s.addShape('roundRect', { x: 2.55, y: 1.5, w: 2.3, h: 4.4, rectRadius: 0.9, fill: { color: '5A5A5A' }, line: NOLINE });
  s.addShape('roundRect', { x: 2.35, y: 2.36, w: 2.7, h: 2.5, rectRadius: 0.5, fill: { color: 'C4C4C4' }, line: NOLINE });
  photo(s, 2.62, 2.62, 2.16, 1.98, { shape: 'roundRect', rectRadius: 0.3 });
  heading(s, [['Best ', C.green], ['Watch', C.white]], 5.855, 1.611, 4.29, { h: 0.774 });
  label(s, 'MOCKUP OF NEWT COMPANY', 5.907, 2.385, 3.018);
  body(s, L181, 5.855, 2.847, 5.007, 0.903);
  statCard(s, 5.545, 4.212, 5.627, 1.504, ['9,5K', '10M', '240']);
  button(s, 'DISCOVER HERE', 1.436, 6.724, 2.411, 0.496, { outline: C.green });
});

// 30 - Consulting Impression (bar chart + laptop)
slides.push(s => {
  backdrop(s);
  chrome(s, '30');
  heading(s, [['Consulting ', C.green], ['Impression', C.white]], 3.142, 1.043, 6.473, { h: 0.774, align: 'center' });
  body(s, L181, 2.33, 2.048, 8.097, 0.625, { align: 'center' });
  s.addChart('bar', [
    { name: 'Series 1', labels: CHART_CATS, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: CHART_CATS, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: CHART_CATS, values: [2, 2, 3, 5] }
  ], Object.assign({ x: 0.86, y: 3.75, w: 5.007, h: 3.169, barDir: 'col', barGapWidthPct: 40 }, CHART_DARK));
  // laptop stand-in
  s.addShape('trapezoid', { x: 8.0, y: 3.588, w: 4.5, h: 2.5, fill: { color: '2E2E2E' }, line: NOLINE, flipV: true });
  photo(s, 8.75, 3.7, 3.1, 2.2, { color: C.photo });
  rect(s, 7.6, 6.05, 5.3, 0.22, 'D2D2D2');
});

// 31 - Best mockup (tilted phones on an accent panel)
slides.push(s => {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  gradBand(s, 0, 0.7, 4.94, 6.8, { vertical: true, steps: 26 });
  rect(s, 0, 0, 13.333, 0.706, C.bar);
  chrome(s, '31');
  [[-0.349, 4.506], [-0.078, 1.616], [2.806, 2.034], [2.59, 4.939]].forEach(([x, y]) => {
    s.addShape('roundRect', { x, y, w: 1.605, h: 2.818, rectRadius: 0.22, fill: { color: C.white }, line: NOLINE, rotate: 51 });
    s.addShape('roundRect', { x: x + 0.1, y: y + 0.18, w: 1.4, h: 2.45, rectRadius: 0.14, fill: { color: C.photo }, line: NOLINE, rotate: 51 });
  });
  heading(s, [['BEST MOCKUP ', C.green], ['NEWT', C.white, F.sans, false], [' ', C.white], ['COMPANY', C.white, F.light]], 6.471, 1.519, 5.106);
  body(s, L196, 6.483, 3.053, 5.365, 0.903);
  [['A', 4.317], ['B', 5.623]].forEach(([n, y]) => {
    bigNum(s, n, 6.52, y + 0.178);
    label(s, 'INSERT YOUR TEXT', 7.361, y, 1.907);
    body(s, L91, 7.361, y + 0.32, 4.374, 0.625, { color: C.gray50 });
  });
});

// 32 - Pricing plans
slides.push(s => {
  backdrop(s);
  chrome(s, '32');
  heading(s, [['PRICING ', C.green], ['PLANS', C.white]], 3.142, 0.846, 6.473, { h: 0.774, align: 'center' });
  body(s, L181, 2.33, 1.851, 8.097, 0.625, { align: 'center' });
  [['BASIC PLANS', '250.35 ', 0.919, 1.12, 1.231, 2.527, 2.27],
   ['STANDARD PLANS', '350.35 ', 4.818, 5.019, 5.13, 6.335, 2.876],
   ['PREMIUM PLANS', '750.35 ', 8.716, 8.917, 9.028, 10.35, 2.764]].forEach(([name, price, cardX, tx, barX, btnX, titleW]) => {
    s.addShape('roundRect', { x: cardX, y: 3.506, w: 3.698, h: 2.575, rectRadius: 0.062 * 2.575, fill: { color: C.bar }, line: NOLINE });
    s.addText(name, { x: tx, y: 3.663, w: titleW, h: 0.572, fontFace: F.xbold, fontSize: 28, bold: true, color: C.white, underline: { style: 'sng' }, valign: 'top' });
    s.addText([
      { text: '$', options: { fontSize: 28, bold: true } },
      { text: price, options: { fontSize: 28, bold: true } },
      { text: '/month', options: { fontSize: 11, bold: true } }
    ], { x: tx, y: 4.179, w: 2.876, h: 0.572, fontFace: F.xbold, color: C.teal, valign: 'top' });
    progress(s, barX, 4.751, 3.147, 0.093, 0.723);
    body(s, L91, tx, 4.954, 3.191, 0.903, { color: C.gray50, align: 'left' });
    button(s, 'PURCHASE NOW', btnX, 5.968, 1.907, 0.5, { radiusFrac: 0.16667, font: F.xbold, bold: false, color: C.offWhite });
  });
});

// 33 - Quote of the day
slides.push(s => {
  backdrop(s);
  s.addShape('ellipse', { x: 4.588, y: 1.871, w: 4.157, h: 4.157, fill: { color: C.bar }, line: NOLINE });
  s.addShape('ellipse', { x: 5.743, y: 3.026, w: 1.848, h: 1.848, fill: { color: ramp(ACCENT, 0.45) }, line: NOLINE });
  chrome(s, '33');
  heading(s, [['QUOTE ', C.green], ['THE DAY', C.white]], 3.142, 0.846, 6.473, { h: 0.774, align: 'center' });
  quoteMark(s, 2.687, 2.552, 0.371, C.green);
  quoteMark(s, 9.861, 5.075, 0.371, C.green);
  s.addText([
    { text: QUOTE_BOLD, options: { bold: true } },
    { text: QUOTE_REST }
  ], { x: 2.308, y: 3.156, w: 8.718, h: 1.717, fontFace: F.sans, fontSize: 24, color: C.white, align: 'center', valign: 'top' });
  button(s, 'NEWT QUOTE', 5.26, 5.791, 2.814, 0.5, { radiusFrac: 0.5, font: F.xbold, bold: false, color: C.offWhite });
  signature(s);
  sideLabel(s, 4.486, 5.595);
});

// 34 - Testimonials
slides.push(s => {
  backdrop(s);
  chrome(s, '34');
  heading(s, [['WHAT ', C.green], ['CLIENT SAYS', C.white]], 3.142, 0.846, 6.473, { h: 0.774, align: 'center' });
  [[1.071, 1.271, 1.383, 2.672], [4.857, 5.056, 5.139, 6.428], [8.642, 8.841, 8.908, 10.198]].forEach(([cardX, imgX, txtX, nameX]) => {
    s.addShape('roundRect', { x: cardX, y: 2.557, w: 3.586, h: 3.186, rectRadius: 0.0977 * 3.186, fill: { color: C.bar }, line: { color: C.green, width: 1 } });
    photo(s, imgX, 2.746, 1.358, 1.358, { shape: 'ellipse' });
    s.addText('Reviewer', { x: nameX, y: 3.11, w: 1.788, h: 0.337, fontFace: F.xbold, fontSize: 14, bold: true, color: C.green, underline: { style: 'sng' }, valign: 'top' });
    label(s, '(80+Review Message)', nameX, 3.447, 1.919, { font: F.light, color: C.white });
    body(s, L108, txtX, 4.314, 3.154, 0.903, { italic: true, align: 'left' });
  });
  signature(s);
});

// 35 - Charts (scatter-style line + wide line chart)
slides.push(s => {
  backdrop(s);
  rect(s, 6.217, 1.369, 6.493, 2.638, C.bar);
  rect(s, 0.812, 4.31, 12.087, 2.638, C.bar);
  chrome(s, '35');
  heading(s, [['Best ', C.green], ['Charts', C.white]], 1.015, 1.611, 4.29, { h: 0.774 });
  label(s, 'CHARTS OF NEWT COMPANY', 1.066, 2.385, 3.018);
  body(s, L181, 1.015, 2.847, 5.007, 0.903);
  s.addChart('line', [
    { name: 'Series 1', labels: CHART_CATS, values: [2.1, 2.0, 3.5, 5.0] },
    { name: 'Series 2', labels: CHART_CATS, values: [2.4, 4.4, 1.9, 2.8] },
    { name: 'Series 3', labels: CHART_CATS, values: [4.3, 2.5, 3.0, 4.5] }
  ], Object.assign({}, CHART_DARK, {
    x: 6.638, y: 1.406, w: 5.449, h: 2.638, showLegend: true, legendPos: 'b',
    lineDataSymbol: 'circle', lineDataSymbolSize: 6, valAxisMaxVal: 6, valAxisMajorUnit: 1
  }));
  s.addChart('line', [
    { name: 'Series 1', labels: CHART_CATS, values: [4.2, 3.0, 2.8, 4.6] },
    { name: 'Series 2', labels: CHART_CATS, values: [6.8, 7.0, 5.4, 7.4] },
    { name: 'Series 3', labels: CHART_CATS, values: [8.8, 9.0, 8.4, 12.4] }
  ], Object.assign({}, CHART_DARK, {
    x: 1.066, y: 4.354, w: 11.543, h: 2.51, showLegend: true, legendPos: 'b',
    lineDataSymbol: 'none', valAxisMaxVal: 14, valAxisMajorUnit: 2
  }));
});

// 36 - Charts (pie + KPI grid)
slides.push(s => {
  backdrop(s);
  chrome(s, '36');
  s.addChart('pie', [{
    name: 'Sales',
    labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    values: [8.2, 3.2, 1.4, 1.2]
  }], Object.assign({}, CHART_DARK, {
    x: 1.212, y: 1.056, w: 4.44, h: 2.96,
    chartColors: [C.teal, C.green, C.gray65, C.gold],
    showLegend: true, legendPos: 'b', showTitle: true, title: 'Sales',
    titleColor: C.gray65, titleFontFace: F.body, titleFontSize: 14,
    dataBorder: { pt: 1, color: C.white }
  }));
  statCard(s, 1.455, 4.203, 3.958, 2.776, ['9,5K', '10M', '2500', '169'], 2);
  heading(s, [['Best ', C.green], ['Charts', C.white]], 6.868, 1.762, 4.29, { h: 0.774 });
  label(s, 'CHARTS OF NEWT COMPANY', 6.919, 2.536, 3.018);
  body(s, L181, 6.868, 2.999, 5.007, 0.903);
  [['A', 4.221], ['B', 5.429]].forEach(([n, y]) => {
    bigNum(s, n, 6.868, y + 0.178);
    label(s, 'INSERT YOUR TEXT', 7.709, y, 1.907);
    body(s, L91, 7.709, y + 0.32, 4.104, 0.625, { color: C.gray50 });
  });
});

// 37 - Infographic (four teardrops)
slides.push(s => {
  backdrop(s);
  chrome(s, '37');
  heading(s, [['INFOGRAPHIC', C.white]], 3.142, 0.846, 6.473, { h: 0.774, align: 'center' });
  s.addShape('teardrop', { x: 4.488, y: 2.356, w: 2.054, h: 2.054, fill: { color: C.gold }, line: NOLINE, rotate: 180, flipH: true });
  s.addShape('teardrop', { x: 6.755, y: 2.356, w: 2.054, h: 2.054, fill: { color: C.teal }, line: NOLINE, rotate: 180 });
  s.addShape('teardrop', { x: 4.488, y: 4.624, w: 2.054, h: 2.054, fill: { color: C.green }, line: NOLINE });
  s.addShape('teardrop', { x: 6.755, y: 4.624, w: 2.054, h: 2.054, fill: { color: C.bar }, line: NOLINE, flipH: true });
  [['A', 5.329, 3.154], ['B', 7.083, 3.154], ['C', 5.329, 4.966], ['D', 7.083, 4.966]].forEach(([n, x, y]) => bigNum(s, n, x, y));
  [[1.118, 1.231, 2.834], [9.137, 9.221, 2.834], [1.118, 1.231, 4.966], [9.137, 9.221, 4.966]].forEach(([tx, dx, y]) => {
    gradBand(s, dx, y - 0.344, 0.303, 0.303, { steps: 5 });
    label(s, 'INSERT YOUR TEXT', tx, y, 1.907);
    body(s, L91, tx, y + 0.32, 2.949, 0.903, { color: C.gray50 });
  });
});

// 38 - Infographic (interlocking rings)
slides.push(s => {
  backdrop(s);
  chrome(s, '38');
  heading(s, [['INFOGRAPHIC', C.green]], 3.142, 0.846, 6.473, { h: 0.774, align: 'center' });
  // Each node = big lettered ring + small ring below, joined by a stem.
  const nodes = [
    ['A', C.green, 4.264, 2.5, 1.24, 4.639, 3.86, 0.66],
    ['B', C.teal, 6.264, 2.5, 1.24, 6.639, 3.86, 0.66],
    ['C', C.gold, 5.264, 4.117, 1.24, 5.639, 3.5, 0.66],
    ['D', C.offWhite, 7.264, 4.117, 1.24, 7.639, 3.5, 0.66]
  ];
  // connecting wave behind the rings
  [[C.green, 3.889, 3.45], [C.green, 5.889, 3.45]].forEach(([col, x, y]) => {
    s.addShape('ellipse', { x, y, w: 1.6, h: 1.05, fill: NOLINE, line: { color: col, width: 3 } });
  });
  nodes.forEach(([letter, col, bx, by, bd, sx, sy, sd]) => {
    s.addShape('ellipse', { x: sx, y: sy, w: sd, h: sd, fill: NOLINE, line: { color: col, width: 3 } });
    s.addShape('rect', { x: (bx + bd / 2) - 0.09, y: by + bd * 0.75, w: 0.18, h: 0.6, fill: { color: col }, line: NOLINE });
    s.addShape('ellipse', { x: bx, y: by, w: bd, h: bd, fill: { color: C.bg }, line: { color: col, width: 3 } });
    s.addText(letter, { x: bx, y: by, w: bd, h: bd, fontFace: F.sans, fontSize: 48, bold: true, color: C.gray88, align: 'center', valign: 'middle' });
  });
  [[1.455, 0.954, 2.78, 3.325], [9.059, 8.585, 2.78, 3.325],
   [2.33, 1.925, 4.983, 5.483], [9.934, 9.556, 4.983, 5.483]].forEach(([bx, tx, by, ty]) => {
    button(s, 'INSERT YOUR TEXT', bx, by, 1.907, 0.5, { radiusFrac: 0.16667, font: F.xbold, bold: false, color: C.offWhite });
    body(s, L91, tx, ty, 2.949, 0.903, { color: C.gray50, align: 'center' });
  });
});

// 39 - Contact
slides.push(s => {
  backdrop(s);
  chrome(s, '39');
  photo(s, 1.471, 1.5, 4.436, 3.574);
  heading(s, [['GET IN TOUCH WITH ', C.green], ['NEWT', C.white, F.sans, false], [' ', C.white], ['COMPANY', C.white, F.light]], 6.219, 1.989, 5.962);
  body(s, L196, 6.231, 3.523, 5.365, 0.903);
  [['816 Greenhills Avenue, New York, CA 71539 , United States', 1.471, 4.747],
   ['www.newt.com', 6.828, 1.811],
   ['+751 656 8654', 9.092, 1.659]].forEach(([text, x, w]) => {
    button(s, text, x, 5.675, w, 0.5, { radiusFrac: 0.16667, font: F.body, size: 11, bold: false });
  });
});

// 40 - Thanks
slides.push(s => {
  rect(s, 0, 0, 13.333, 7.5, C.bg);
  s.addShape('parallelogram', { x: 6.796, y: 0, w: 5.81, h: 7.5, fill: { color: C.panel }, line: NOLINE });
  s.addShape('parallelogram', { x: 0.58, y: 0, w: 5.81, h: 7.5, fill: { color: C.panel }, line: NOLINE, flipH: true });
  gradRing(s, 3.909, 0.992, 5.515);
  s.addShape('ellipse', { x: 4.048, y: 1.131, w: 5.238, h: 5.238, fill: { color: C.bar }, line: NOLINE });
  s.addText('THANKS', { x: 3.049, y: 2.683, w: 7.236, h: 1.447, fontFace: F.sans, fontSize: 80, bold: true, color: C.white, align: 'center', valign: 'top' });
  body(s, L142, 4.246, 4.131, 4.841, 0.903, { align: 'center' });
  gradPill(s, 4.464, 5.961, 4.406, 0.711, 0.43158 * 0.711);
  s.addText('Thank you for watching us', { x: 4.654, y: 6.089, w: 4.017, h: 0.438, fontFace: F.xbold, fontSize: 20, color: C.offWhite, valign: 'top' });
});

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
const pptx = new pptxgen();
pptx.defineLayout({ name: 'NEWT', width: 13.333, height: 7.5 });
pptx.layout = 'NEWT';
pptx.title = 'NEWT Company Presentation';
pptx.author = 'Yumnacreative';

slides.forEach(build => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '157da01b-e355-43ba-9e76-91cea83065b6_grok_final.pptx') })
  .then(f => console.log('wrote', f));
