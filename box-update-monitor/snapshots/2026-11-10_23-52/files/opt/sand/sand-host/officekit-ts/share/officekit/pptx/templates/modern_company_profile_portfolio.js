/**
 * MOSY — minimal portrait presentation template (43 slides, 17.22" x 24.36").
 * Rebuilt with pptxgenjs. Photographs in the original are picture placeholders;
 * they are recreated here as light "[image]" panels.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const W = 17.2222, H = 24.3611;          // slide size, inches
const INK = '111111';                    // primary text
const GRAY = '5E5E5E';                   // secondary text
const BLACK = '000000';
const TEAL = '637A7E';                   // blob colour A
const GOLD = 'E4B35A';                   // blob colour B
const PANEL = 'F3F4F7';                  // photo placeholder fill
const PANEL_TX = 'C4C9D2';               // photo placeholder caption
const DEV_DARK = '24242D';               // device body / screen
const DEV_EDGE = '151314';               // device outer frame
const DEV_LIGHT = 'DCDEE0';              // light device band

/* ------------------------------------------------------------------- typography */
const F_DISPLAY = 'Montserrat Black';
const F_HEAD = 'Open Sans';
const F_BODY = 'Open Sans Light';

// Text presets used throughout the deck.
const DISPLAY = { fontFace: F_DISPLAY, fontSize: 96, color: INK, charSpacing: 4.8, lineSpacingMultiple: 0.9 };
const NUMBER = { fontFace: F_DISPLAY, fontSize: 110, color: INK, charSpacing: 5.5, lineSpacingMultiple: 0.8 };
const HEAD = { fontFace: F_HEAD, fontSize: 42, bold: true, color: INK, charSpacing: 0.84, lineSpacingMultiple: 0.9 };
const BODY = { fontFace: F_BODY, fontSize: 30, color: INK, lineSpacingMultiple: 1.2 };
const BODY_GRAY = Object.assign({}, BODY, { color: GRAY });
const LABEL = Object.assign({}, BODY_GRAY, { valign: 'middle' });
const DROP = { fontFace: 'Montserrat Semi Bold', fontSize: 60, color: INK, wrap: false, margin: 7.2 };

/* ---------------------------------------------------------------- primitives */
// Add a text box. `lines` is a string or an array of strings (one paragraph each).
function txt(s, lines, x, y, w, h, preset, extra) {
  const o = Object.assign({ x, y, w, h, margin: 3.6, align: 'left', valign: 'top' }, preset, extra);
  const arr = Array.isArray(lines) ? lines : [lines];
  const runs = arr.map((t, i) => ({
    text: t,
    options: { breakLine: i < arr.length - 1, bullet: o.bullet, indentLevel: o.indentLevel },
  }));
  s.addText(runs, o);
}

// Photo placeholder: light panel standing in for a photograph.
function photo(s, x, y, w, h, opts) {
  const o = opts || {};
  s.addShape(o.round ? 'roundRect' : 'rect',
    Object.assign({ x, y, w, h, fill: { color: PANEL } }, o.round ? { rectRadius: Math.min(w, h) * 0.12 } : {}));
  if (w > 1 && h > 0.6) {
    s.addText('[image]', {
      x, y: y + h / 2 - 0.35, w, h: 0.7, align: 'center', valign: 'middle',
      fontFace: F_BODY, fontSize: 20, color: PANEL_TX,
    });
  }
}

// The "- Drop Image here" caption that the template prints over its mockups.
function dropCaption(s, x, y, w, h) {
  txt(s, ['', '- Drop Image here'], x, y, w, h, DROP);
}

/* --------------------------------------------------------------------- blobs */
// Organic background shapes. Each entry is a closed cubic path in unit space:
// [startX, startY] followed by [c1x,c1y, c2x,c2y, x,y] segments.
const BLOBS = {
  pebble: [[0.968, 0.409],
    [1.049, 0.597, 0.969, 0.825, 0.810, 0.932], [0.581, 1.085, 0.302, 0.965, 0.127, 0.716],
    [0.054, 0.612, 0.000, 0.487, 0.000, 0.351], [0.000, 0.234, 0.041, 0.121, 0.121, 0.056],
    [0.284, -0.076, 0.487, 0.055, 0.668, 0.153], [0.785, 0.217, 0.906, 0.266, 0.968, 0.409]],
  kidney: [[0.871, 0.198],
    [1.043, 0.381, 1.043, 0.619, 0.871, 0.802], [0.604, 1.085, 0.068, 1.046, 0.005, 0.802],
    [-0.023, 0.695, 0.080, 0.601, 0.080, 0.500], [0.080, 0.399, -0.023, 0.305, 0.005, 0.198],
    [0.068, -0.046, 0.604, -0.085, 0.871, 0.198]],
  drop: [[0.976, 0.346],
    [1.043, 0.574, 0.966, 0.832, 0.802, 0.943], [0.586, 1.090, 0.341, 0.932, 0.152, 0.709],
    [0.061, 0.600, -0.023, 0.460, 0.006, 0.304], [0.034, 0.148, 0.157, 0.077, 0.276, 0.039],
    [0.560, -0.052, 0.874, 0.003, 0.976, 0.346]],
  egg: [[0.978, 0.343],
    [1.042, 0.566, 0.963, 0.814, 0.804, 0.931], [0.699, 1.008, 0.584, 1.017, 0.472, 0.976],
    [0.361, 0.934, 0.255, 0.844, 0.163, 0.737], [0.062, 0.621, -0.028, 0.468, 0.008, 0.302],
    [0.041, 0.151, 0.161, 0.081, 0.278, 0.042], [0.564, -0.054, 0.879, 0.003, 0.978, 0.343]],
  wave: [[0.963, 0.239],
    [1.052, 0.486, 0.973, 0.801, 0.809, 0.933], [0.587, 1.110, 0.340, 0.909, 0.143, 0.646],
    [0.044, 0.515, -0.039, 0.329, 0.019, 0.163], [0.072, 0.008, 0.198, 0.025, 0.309, 0.025],
    [0.555, 0.026, 0.838, -0.109, 0.963, 0.239]],
};

function blob(s, name, x, y, w, h, color, rotate, flipH) {
  const p = BLOBS[name];
  const pts = [{ x: p[0][0] * w, y: p[0][1] * h, moveTo: true }];
  for (let i = 1; i < p.length; i++) {
    const c = p[i];
    pts.push({ x: c[4] * w, y: c[5] * h, curve: { type: 'cubic', x1: c[0] * w, y1: c[1] * h, x2: c[2] * w, y2: c[3] * h } });
  }
  pts.push({ close: true });
  s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color }, rotate: rotate || 0, flipH: !!flipH });
}

/* --------------------------------------------------------------------- icons */
// Circle with a check mark.
function iconCheck(s, x, y, d, color) {
  s.addShape('ellipse', { x, y, w: d, h: d, line: { color, width: 1.5 } });
  s.addShape('line', { x: x + d * 0.28, y: y + d * 0.50, w: d * 0.16, h: d * 0.18, line: { color, width: 1.5 } });
  s.addShape('line', { x: x + d * 0.44, y: y + d * 0.32, w: d * 0.28, h: d * 0.36, flipV: true, line: { color, width: 1.5 } });
}

// Flat social marks used next to profiles and follower counts.
function iconSocial(s, kind, x, y, w, h, color) {
  if (kind === 'instagram') {
    s.addShape('roundRect', { x, y, w, h, rectRadius: w * 0.28, line: { color, width: 2.5 } });
    s.addShape('ellipse', { x: x + w * 0.27, y: y + h * 0.27, w: w * 0.46, h: h * 0.46, line: { color, width: 2.5 } });
    s.addShape('ellipse', { x: x + w * 0.72, y: y + h * 0.14, w: w * 0.1, h: h * 0.1, fill: { color } });
  } else if (kind === 'youtube') {
    s.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(w, h) * 0.3, fill: { color } });
    s.addShape('triangle', { x: x + w * 0.38, y: y + h * 0.28, w: w * 0.3, h: h * 0.44, rotate: 90, fill: { color: 'FFFFFF' } });
  } else if (kind === 'pinterest') {
    s.addShape('ellipse', { x, y, w, h, fill: { color } });
    txt(s, 'p', x, y - h * 0.06, w, h, { fontFace: 'Georgia', fontSize: h * 62, bold: true, color: 'FFFFFF' },
      { align: 'center', valign: 'middle' });
  } else if (kind === 'medium') {
    s.addShape('rect', { x, y, w, h, fill: { color } });
    txt(s, 'M', x, y, w, h, { fontFace: 'Georgia', fontSize: h * 46, color: 'FFFFFF' }, { align: 'center', valign: 'middle' });
  } else if (kind === 'twitter') {
    s.addShape('moon', { x: x - w * 0.15, y, w: w * 1.15, h, rotate: 205, fill: { color } });
    s.addShape('triangle', { x: x + w * 0.1, y: y + h * 0.3, w: w * 0.7, h: h * 0.5, rotate: 200, fill: { color } });
  }
}

// Paper plane (slide 5): outlined body plus the folded-wing crease.
function iconPlane(s, x, y, d, color) {
  const ln = { color, width: 3.5 };
  const poly = pts => pts.map((p, i) => ({ x: p[0] * d, y: p[1] * d, moveTo: i === 0 }))
    .concat([{ close: true }]);
  s.addShape('custGeom', { x, y, w: d, h: d, line: ln, points: poly([[0.00, 0.44], [0.98, 0.05], [0.60, 1.00], [0.40, 0.68]]) });
  s.addShape('custGeom', { x, y, w: d, h: d, line: ln, points: poly([[0.98, 0.05], [0.40, 0.68], [0.30, 0.92]]) });
}

/* ------------------------------------------------------------------- devices */
// Phone / tablet / watch / laptop mock-ups drawn from plain shapes.
function phoneMock(s, x, y, w, h, screen, landscape) {
  const m = Math.min(w, h) * 0.045;
  s.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(w, h) * 0.16, fill: { color: DEV_EDGE } });
  s.addShape('roundRect', {
    x: x + m, y: y + m, w: w - 2 * m, h: h - 2 * m,
    rectRadius: Math.min(w, h) * 0.13, fill: { color: screen },
  });
  // Front-camera notch: top edge when upright, left edge when laid on its side.
  s.addShape('roundRect', landscape
    ? { x: x + m, y: y + h * 0.34, w: w * 0.035, h: h * 0.32, rectRadius: 0.03, fill: { color: DEV_EDGE } }
    : { x: x + w * 0.34, y: y + m, w: w * 0.32, h: h * 0.035, rectRadius: 0.03, fill: { color: DEV_EDGE } });
}

function tabletMock(s, x, y, w, h, screen) {
  s.addShape('roundRect', { x, y, w, h, rectRadius: Math.min(w, h) * 0.07, fill: { color: DEV_EDGE } });
  const m = Math.min(w, h) * 0.035;
  s.addShape('rect', { x: x + m, y: y + m, w: w - 2 * m, h: h - 2 * m, fill: { color: screen } });
}

function watchMock(s, x, y, w, h, screen, band) {
  s.addShape('roundRect', { x: x + w * 0.16, y, w: w * 0.68, h: h * 0.28, rectRadius: 0.06, fill: { color: band } });
  s.addShape('roundRect', { x: x + w * 0.16, y: y + h * 0.72, w: w * 0.68, h: h * 0.28, rectRadius: 0.06, fill: { color: band } });
  s.addShape('roundRect', { x, y: y + h * 0.17, w, h: h * 0.66, rectRadius: w * 0.24, fill: { color: '6D6C6F' } });
  s.addShape('roundRect', { x: x + w * 0.05, y: y + h * 0.20, w: w * 0.90, h: h * 0.60, rectRadius: w * 0.20, fill: { color: DEV_EDGE } });
  s.addShape('roundRect', { x: x + w * 0.09, y: y + h * 0.23, w: w * 0.82, h: h * 0.54, rectRadius: w * 0.17, fill: { color: screen } });
  s.addShape('roundRect', { x: x + w * 0.955, y: y + h * 0.34, w: w * 0.055, h: h * 0.11, rectRadius: 0.02, fill: { color: '6D6C6F' } });
}

function laptopMock(s, x, y, w, h, screen) {
  const lidW = w * 0.82, lidH = h * 0.93, lidX = x + (w - lidW) / 2;
  s.addShape('roundRect', { x: lidX, y, w: lidW, h: lidH, rectRadius: 0.06, fill: { color: DEV_EDGE } });
  s.addShape('rect', { x: lidX + w * 0.012, y: y + h * 0.02, w: lidW - w * 0.024, h: lidH - h * 0.04, fill: { color: screen } });
  s.addShape('roundRect', { x, y: y + lidH, w, h: h * 0.045, rectRadius: 0.03, fill: { color: '9A9DA2' } });
}

/* ----------------------------------------------------------- icon sheet grid */
// Slides 40-43 are bonus sheets of line icons. Each icon is approximated with an
// outlined preset shape drawn on the same 12-column grid used by the original.
const ICON_SETS = {
  general: ['lightningBolt', 'triangle', 'teardrop', 'gear6', 'gear9', 'rect', 'wedgeRoundRectCallout', 'homePlate',
    'home', 'pentagon', 'flowChartProcess', 'plaque', 'ellipse', 'trapezoid', 'ellipse', 'donut',
    'flowChartDocument', 'roundRect', 'hexagon', 'octagon', 'mathPlus', 'mathMinus', 'chord', 'mathMultiply',
    'mathDivide', 'noSmoking', 'cloud', 'arc', 'frame', 'halfFrame', 'bevel', 'can', 'cube', 'diamond',
    'chevron', 'corner', 'parallelogram', 'pie', 'blockArc', 'flowChartConnector'],
  place: ['ellipse', 'teardrop', 'donut', 'pie', 'flowChartConnector', 'blockArc', 'arc', 'chord', 'moon', 'sun',
    'star5', 'triangle'],
  design: ['line', 'parallelogram', 'rtTriangle', 'rect', 'frame', 'halfFrame', 'trapezoid', 'can', 'cube',
    'flowChartDocument', 'diamond', 'plaque'],
  weather: ['sun', 'ellipse', 'moon', 'cloud', 'cloud', 'cloud', 'cloud', 'cloud', 'star4'],
  message: ['rect', 'roundRect', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'cloudCallout', 'ellipse',
    'flowChartDelay', 'callout1', 'borderCallout1', 'wedgeEllipseCallout'],
  people: ['ellipse', 'smileyFace', 'ellipse', 'flowChartConnector', 'ellipse', 'smileyFace'],
  smiley: ['smileyFace'],
  arrows: ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'bentArrow', 'bentUpArrow', 'curvedRightArrow',
    'curvedLeftArrow', 'uturnArrow', 'circularArrow', 'leftRightArrow', 'upDownArrow', 'notchedRightArrow',
    'stripedRightArrow', 'quadArrow', 'swooshArrow', 'chevron', 'leftUpArrow'],
  shop: ['flowChartInputOutput', 'rect', 'roundRect', 'can', 'flowChartMagneticDisk', 'cube', 'diamond',
    'flowChartDocument', 'plaque', 'trapezoid', 'hexagon', 'chartPlus'],
  devices: ['rect', 'roundRect', 'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDisplay',
    'flowChartInternalStorage', 'flowChartManualInput', 'snip2SameRect', 'round2SameRect', 'frame'],
  circle: ['ellipse'],
  socialGlyph: ['ellipse', 'roundRect', 'flowChartConnector', 'teardrop', 'rect', 'pentagon'],
};

// One row of icons: `n` shapes at `y`, stepping 1.205" across the sheet.
function iconRow(s, x0, y, n, set, offset, filled) {
  const pitch = 1.205, d = 0.47;
  for (let i = 0; i < n; i++) {
    const shape = set[(i + offset) % set.length];
    const o = { x: x0 + i * pitch, y, w: d, h: d };
    s.addShape(shape, filled ? Object.assign(o, { fill: { color: BLACK } })
      : Object.assign(o, { line: { color: INK, width: 1.25 } }));
  }
}

function iconSheet(s, x0, sections) {
  sections.forEach(sec => {
    if (sec.label) txt(s, sec.label, sec.labelX || 1.389, sec.labelY, 5.5, 0.677, LABEL);
    sec.rows.forEach((row, r) => iconRow(s, x0, row[0] - 0.235, row[1], ICON_SETS[sec.set], r * 5, sec.filled));
  });
}

/* ------------------------------------------------------------ slide builders */
const build = [];

// 1 — read-me / instructions page
build.push(s => {
  txt(s, 'Check before you start', 1.398, 1.917, 13.788, 0.856, HEAD);
  txt(s, [
    '1. Fonts: If you have the font that was used in this template on your pc this will be displayed correctly.',
    'If the fonts NOT displayed correctly: Download the "fonts" we use that you see the same style as in the preview! ',
    '2. Photos: Its never allowed to add third parties images/files to a download file.',
    'So check our Unsplash & Graphicgum Collection with used images in preview, download it for free and place it in!',
    '3. Our Bonus Vector Line Icons you find on the last slides. Easy to edit in color and size.',
  ], 1.398, 4.164, 14.426, 6.819, BODY, { bullet: { characterCode: '002D', indent: 28.8 } });
  txt(s, '1. Install: Open Sans and Montserrat (Check link or other Website were you can download it)',
    1.398, 12.139, 14.426, 1.353, BODY, { bullet: { characterCode: '002D', indent: 28.8 } });
  const link = { fontFace: F_HEAD, fontSize: 28, color: '0B0C11', underline: { style: 'sng' }, lineSpacingMultiple: 1.2, valign: 'middle' };
  txt(s, 'https://www.fontsquirrel.com/fonts/open-sans', 1.419, 14.335, 10.650, 0.720, link);
  txt(s, 'https://www.fontsquirrel.com/fonts/montserrat', 1.419, 15.257, 10.650, 0.720, link);
  txt(s, '2. Image Collection „check and download“ Images that used in this template.',
    1.398, 16.935, 14.426, 1.325, BODY, { bullet: { characterCode: '002D', indent: 28.8 } });
  txt(s, 'https://unsplash.com/collections/3856769/', 1.419, 18.637, 14.426, 0.668, link);
  txt(s, 'https:/graphicgum.com', 1.419, 19.682, 14.544, 0.668,
    { fontFace: 'Montserrat', fontSize: 26, color: INK, underline: { style: 'sng' }, lineSpacingMultiple: 1.2 });
  txt(s, 'For more Info and help please read the „Documentation“ in your download file. Pixasquare says thanks!',
    1.419, 20.983, 14.426, 1.344, BODY);
});

// 2 — cover
build.push(s => {
  photo(s, 1.398, 1.392, 8.667, 10.833);
  txt(s, 'this is', 1.398, 13.509, 6.889, 1.711, DISPLAY);
  txt(s, 'modern clean and minimal design.', 1.398, 19.980, 5.694, 1.604, HEAD);
  txt(s, 'By Pixasquare', 12.999, 20.904, 2.847, 0.681, BODY);
  blob(s, 'pebble', 8.061, 11.925, 5.902, 4.877, TEAL, 36.87);
  txt(s, 'MOSY', 10.067, 13.509, 6.103, 1.711, DISPLAY);
});

// 3 — table of contents
build.push(s => {
  blob(s, 'kidney', 1.574, 17.778, 2.878, 4.125, GOLD, 142.03);
  ['1. The Company', '2. About us', '3. Our Creative Team', '4. Goals & Vision', '5. Projects & Partners',
    '6. Collection - Portfolio', '7. Marketing & Service', '8. We growth your Business', '9. Statistics - Markets',
    '10. Contact - Social Media',
  ].forEach((line, i) => txt(s, line, 1.398, 4.164 + i * 0.8976, 5.833, 0.669, BODY, { valign: 'middle' }));
  txt(s, 'content', 1.398, 19.568, 7.278, 1.711, DISPLAY);
  photo(s, 10.484, 4.163, 5.362, 9.410);
});

// 4 — about
build.push(s => {
  photo(s, 8.889, 9.324, 8.333, 15.037);
  txt(s, 'about', 1.398, 2.690, 7.095, 1.711, DISPLAY);
  photo(s, 1.398, 6.467, 3.775, 2.857);
  blob(s, 'kidney', 5.384, 11.327, 4.899, 7.021, GOLD, 319.39);
  txt(s, 'Maorem eso ipsum dolor sito amet, derasm the consetetur sadipdniec nonumy eirmod tempor invidunt on utony them etodre indsa dolore.',
    1.398, 14.034, 6.931, 2.731, BODY);
  txt(s, 'Modern clean and minimal design.', 1.398, 19.980, 5.694, 1.604, HEAD);
});

// 5 — our value
build.push(s => {
  photo(s, -0.021, 0, 10.858, 13.573);
  txt(s, 'our value', 1.398, 8.808, 4.915, 3.161, DISPLAY, { valign: 'middle' });
  txt(s, 'Modern clean and minimal design.', 1.398, 15.562, 5.694, 1.604, HEAD, { valign: 'middle' });
  const lorem = 'Maorem eso ipsum dolor sito amet, derasm the consetetur sadipsing elitromas. diamy utony them labore etodre dolore.';
  txt(s, lorem, 8.976, 15.562, 6.870, 2.731, BODY);
  txt(s, lorem, 8.976, 18.854, 6.870, 2.731, BODY);
  blob(s, 'kidney', 3.140, 19.675, 1.453, 2.083, TEAL, 40.6);
  iconPlane(s, 2.774, 19.884, 1.333, BLACK);
});

// 6 — vision
build.push(s => {
  photo(s, 5.222, 0, 12.0, 15.0);
  blob(s, 'drop', 2.428, 4.158, 5.606, 4.152, GOLD, 117.3);
  txt(s, 'vision', 0.917, 4.972, 6.192, 1.711, DISPLAY);
  txt(s, 'Modern clean and minimal design.', 1.398, 16.125, 5.694, 1.604, HEAD);
  txt(s, 'Maorem eso ipsum dolor sito amet, derasm the consetetur sadipsing elitromas. diamy utony them labore etodre dolore.',
    1.398, 18.854, 6.972, 2.731, BODY);
});

// 7 — our goal
build.push(s => {
  photo(s, 5.222, 9.361, 12.0, 15.0);
  txt(s, 'this year.', 1.398, 2.774, 5.694, 0.892, HEAD);
  txt(s, 'Demas eso ipsum dolor sito amet, derasm the consetetur sadipsing elitromas. diamy utony them lam etodre imas dolore.',
    1.398, 4.364, 6.972, 2.731, BODY);
  txt(s, 'our goal', 11.209, 4.006, 4.637, 3.161, DISPLAY);
  blob(s, 'egg', 1.926, 15.277, 4.878, 3.671, TEAL, 117.3);
  txt(s, '45%', 2.046, 14.944, 4.637, 1.711, DISPLAY);
});

// 8 — team profile
build.push(s => {
  photo(s, 1.398, 1.392, 6.528, 12.181);
  txt(s, ['lora', 'toms'], 10.753, 6.632, 5.093, 3.161, DISPLAY);
  iconSocial(s, 'instagram', 10.988, 11.736, 0.444, 0.444, INK);
  iconSocial(s, 'pinterest', 11.785, 11.737, 0.444, 0.443, INK);
  iconSocial(s, 'youtube', 12.554, 11.783, 0.500, 0.350, INK);
  iconSocial(s, 'medium', 13.379, 11.735, 0.444, 0.444, INK);
  txt(s, 'Stylist & Fashion', 1.398, 15.852, 5.972, 0.892, HEAD);
  txt(s, 'Dons mpor invidunt utony thenoy labore etodre lore desando frond iron laresca in this greadon.',
    1.398, 17.701, 5.359, 2.731, BODY);
  [['2010-2015', 'Master of Art'], ['2015-2018', 'CEO  Rawbelow'], ['2018-Now', 'Mosy Fashion']]
    .forEach((row, i) => {
      txt(s, row[0], 9.580, 17.701 + i * 0.8125, 2.415, 0.681, BODY);
      txt(s, row[1], 12.677, 17.701 + i * 0.8125, 3.236, 0.681, BODY);
    });
});

// 9 — our team
build.push(s => {
  blob(s, 'kidney', 2.423, 2.901, 2.899, 4.154, GOLD, 40.6);
  txt(s, ['our', 'team'], 1.398, 2.774, 6.123, 3.161, DISPLAY);
  [1.398, 6.595, 11.793].forEach((x, i) => {
    photo(s, x, 9.344, i === 2 ? 4.032 : 4.132, 6.528);
    txt(s, 'JONA MILLER', x, 16.425, 4.132, 0.669, BODY);
    txt(s, 'Desndo frond idosan lares this greadon.', x, 17.382, 4.132, 1.364, BODY);
  });
});

// 10 — art director
build.push(s => {
  photo(s, 1.443, 9.721, 8.333, 10.829);
  iconSocial(s, 'instagram', 1.443, 4.165, 0.444, 0.444, INK);
  iconSocial(s, 'pinterest', 2.240, 4.166, 0.444, 0.443, INK);
  iconSocial(s, 'youtube', 3.009, 4.212, 0.500, 0.350, INK);
  iconSocial(s, 'medium', 3.833, 4.164, 0.444, 0.444, INK);
  txt(s, 'Art Director', 1.443, 5.354, 5.972, 0.892, HEAD);
  txt(s, 'Dons mpor invidunt utony ons thenoy labore etodre lore mera desando frond iron laresca in this donsa merec.',
    9.784, 4.166, 6.017, 2.731, BODY);
  txt(s, 'Esando frond iron laresca inos this oras coera.', 9.829, 7.289, 6.017, 1.364, BODY);
  blob(s, 'egg', 7.338, 15.403, 4.878, 3.671, TEAL, 117.3);
  txt(s, ['moki', 'oka'], 9.427, 15.740, 6.139, 3.161, DISPLAY);
});

// 11 — service, four numbered blobs
build.push(s => {
  const items = [
    { n: '1', bx: 2.662, by: 8.065, bw: 1.610, bh: 1.192, blobName: 'drop', bc: GOLD, br: 117.3, nx: 2.774, ny: 7.225, tx: 4.825, ty: 7.686 },
    { n: '2', bx: 2.742, by: 11.788, bw: 1.618, bh: 1.337, blobName: 'pebble', bc: TEAL, br: 194.04, nx: 3.029, ny: 11.015, tx: 5.080, ty: 11.569 },
    { n: '3', bx: 2.665, by: 15.432, bw: 2.028, bh: 1.234, blobName: 'wave', bc: GOLD, br: 117.3, nx: 2.778, ny: 14.875, tx: 5.103, ty: 15.429 },
    { n: '4', bx: 2.704, by: 19.417, bw: 1.742, bh: 1.439, blobName: 'pebble', bc: TEAL, br: 117.3, nx: 2.778, ny: 18.661, tx: 5.103, ty: 18.985 },
  ];
  items.forEach(i => blob(s, i.blobName, i.bx, i.by, i.bw, i.bh, i.bc, i.br));
  txt(s, 'service', 1.398, 2.231, 7.734, 1.711, DISPLAY);
  photo(s, 9.132, 4.164, 3.056, 2.360);
  photo(s, 12.784, 4.164, 3.056, 2.360);
  items.forEach(i => {
    txt(s, i.n, i.nx, i.ny, 1.389, 1.933, NUMBER, { align: 'center' });
    txt(s, 'Onsa the clita kuargren, nosotrad seat the os takimta sapsusit drino for the druamet comda.',
      i.tx, i.ty, 6.708, 2.036, BODY);
  });
});

// 12 — projects
build.push(s => {
  photo(s, 1.271, 6.584, 7.500, 7.500);
  photo(s, 1.271, 14.084, 7.500, 7.500);
  txt(s, 'Rmod tempor idosn invidunt utony them labore etodre dolore.', 1.398, 2.774, 4.305, 2.047, BODY);
  txt(s, 'Modern clean and minimal design.', 10.152, 6.584, 5.694, 1.604, HEAD, { valign: 'middle' });
  txt(s, 'Maorem eso ipsum dolor sito amet, derasm the consetetur sadipscing elitromas. Sedma diam design by pixasquare tempor invidunt utony them labore etodre dolore.',
    10.152, 9.987, 5.694, 4.097, BODY);
  blob(s, 'pebble', 6.986, 17.587, 3.365, 2.781, TEAL, 194.04);
  txt(s, 'projects', 8.293, 17.679, 7.553, 1.711, DISPLAY, { valign: 'middle' });
});

// 13 — partners
build.push(s => {
  blob(s, 'wave', 10.028, 11.650, 6.050, 3.681, TEAL, 330.95);
  blob(s, 'pebble', 2.095, 16.903, 5.422, 4.480, GOLD, 117.3);
  txt(s, 'partners', 1.398, 2.774, 8.134, 1.711, DISPLAY);
  photo(s, 1.398, 6.527, 8.134, 4.533);
  txt(s, 'Do notrad seat the os indsa takimta sancem ipsusit drin for amet comda murma.', 10.793, 13.098, 4.167, 2.719, BODY);
  photo(s, 4.124, 14.748, 3.056, 0.778);
  txt(s, 'Do notrad seat the os design by pixasquare sancem ipsusit drin for amet sons undas incs comda.',
    4.124, 17.588, 4.167, 3.403, BODY);
  photo(s, 10.793, 18.167, 3.290, 0.780);
});

// 14 — new collection
build.push(s => {
  txt(s, 'new collection', 1.398, 2.885, 16.086, 1.711, DISPLAY);
  photo(s, 1.398, 7.191, 5.362, 6.702);
  photo(s, 6.760, 7.191, 5.362, 6.702);
  photo(s, 12.121, 13.893, 5.101, 6.702);
  txt(s, 'Uno etoa jaom froan modern clean creative design.', 1.398, 18.991, 9.192, 1.604, HEAD, { valign: 'middle' });
});

// 15 — portfolio (left copy)
build.push(s => {
  blob(s, 'pebble', 1.262, 2.960, 2.629, 2.172, GOLD, 20.56);
  txt(s, 'new style', 1.817, 3.553, 3.195, 0.892, HEAD);
  txt(s, 'portfolio', 1.398, 9.361, 8.720, 1.711, DISPLAY, { valign: 'middle' });
  photo(s, 11.222, 9.361, 6.0, 7.500);
  photo(s, 11.222, 16.861, 6.0, 7.500);
  txt(s, 'Modern clean and minimal design.', 1.398, 15.257, 5.694, 1.604, HEAD);
  txt(s, 'Cons ustur sadipscing inasm elitromas. Sedma diam dniec nonumy eirmod tempor don invidunt utony them labore etodre dolore.',
    1.398, 18.171, 5.694, 3.414, BODY);
});

// 16 — portfolio (grid)
build.push(s => {
  blob(s, 'wave', 9.640, 2.491, 6.453, 3.926, TEAL, 330.95);
  txt(s, 'portfolio', 7.640, 2.709, 8.206, 1.711, DISPLAY, { valign: 'middle' });
  photo(s, 0, 9.361, 10.105, 9.691);
  photo(s, 10.496, 9.361, 6.726, 7.157);
  photo(s, 10.496, 16.861, 6.726, 7.500);
  photo(s, 0.007, 19.443, 10.098, 4.918);
});

// 17 — the style
build.push(s => {
  txt(s, ['new', 'minimal design.'], 1.398, 1.848, 4.083, 2.317, HEAD, { valign: 'middle' });
  txt(s, ['the', 'style'], 1.398, 5.781, 4.817, 3.161, DISPLAY, { valign: 'middle' });
  photo(s, 8.540, 5.744, 8.682, 5.646);
  [0, 5.835, 11.671].forEach(x => photo(s, x, 11.565, 5.623, 5.623));
  photo(s, 8.335, 17.401, 5.646, 5.646);
});

// 18 — our app (phone mock-up)
build.push(s => {
  blob(s, 'drop', 1.581, 2.895, 1.610, 1.192, GOLD, 21.8);
  txt(s, 'new', 2.069, 2.833, 1.805, 0.892, HEAD, { valign: 'middle' });
  txt(s, 'our app', 9.765, 4.164, 6.081, 1.711, DISPLAY);
  phoneMock(s, 10.372, 8.083, 5.351, 10.827, PANEL);
  photo(s, 10.670, 8.292, 4.869, 10.470, { round: true });
  dropCaption(s, 10.670, 8.292, 4.869, 10.470);
  [9.541, 12.513, 15.569].forEach(y => {
    iconCheck(s, 1.417, y, 0.889, INK);
    txt(s, 'Nosotrad on seat the osno takimta sancem ipsusit drin for the druamet comda.', 3.250, y + 0.025, 5.361, 2.036, BODY);
  });
});

// 19 — online shop
build.push(s => {
  blob(s, 'pebble', 0.993, 6.590, 2.270, 1.876, TEAL, 19.65);
  txt(s, 'Dniec numy eirmod ondas es tempoa utony them labore etodre dolore.', 1.426, 2.774, 7.368, 1.364, BODY);
  txt(s, 'online', 1.343, 6.092, 6.179, 1.711, DISPLAY);
  txt(s, 'www.company.com', 1.426, 10.496, 7.368, 0.892, HEAD);
  txt(s, 'our new online shop style', 1.426, 13.573, 5.646, 1.604, HEAD);
  photo(s, 1.426, 15.938, 5.646, 5.646);
  photo(s, 10.150, 10.724, 5.695, 10.860);
});

// 20 — new style
build.push(s => {
  photo(s, 0, 1.392, 8.611, 12.917);
  photo(s, 8.611, 10.056, 8.611, 12.917);
  txt(s, ['new', 'style'], 1.389, 18.423, 4.626, 3.161, DISPLAY);
});

// 21 — look book (tablet mock-up)
build.push(s => {
  txt(s, ['look', 'book'], 1.398, 6.510, 6.123, 3.161, DISPLAY);
  iconCheck(s, 9.435, 8.450, 0.633, BLACK);
  txt(s, 'Dm indos takimta sanctus esto ipsusit drin for the the nuras murda insam ons druamet.', 10.808, 8.311, 5.038, 2.719, BODY);
  tabletMock(s, 3.007, 13.417, 11.212, 8.040, PANEL);
  photo(s, 3.194, 13.640, 10.833, 7.628, { round: true });
  dropCaption(s, 3.194, 13.640, 10.833, 7.628);
});

// 22 — social media plan
build.push(s => {
  blob(s, 'pebble', 1.258, 1.599, 4.094, 3.383, TEAL, 14.04);
  txt(s, 'plan ', 2.903, 2.518, 4.736, 1.711, DISPLAY);
  txt(s, ['social', 'media'], 7.388, 6.014, 5.625, 3.161, DISPLAY);
  photo(s, 1.398, 10.751, 2.813, 6.082);
  const rows = [
    { icon: 'instagram', ix: 7.416, iy: 10.959, iw: 0.889, ih: 0.889, cx: 7.388, count: '219k', cy: 12.554, ty: 10.751, copy: 'desm' },
    { icon: 'youtube', ix: 7.444, iy: 14.864, iw: 0.889, ih: 0.623, cx: 7.444, count: '84k', cy: 16.387, ty: 14.559, copy: 'desm' },
    { icon: 'pinterest', ix: 7.468, iy: 19.073, iw: 0.893, ih: 0.889, cx: 7.468, count: '17k', cy: 20.693, ty: 18.865, copy: 'desn' },
  ];
  rows.forEach(r => {
    iconSocial(s, r.icon, r.ix, r.iy, r.iw, r.ih, BLACK);
    txt(s, r.count, r.cx, r.cy, 1.833, 0.892, HEAD, { valign: 'middle' });
    txt(s, 'Deot clita kuasd conm  ' + r.copy + ' estasmae, notrad seat the os indsa takimta sancem ipsusit drin for amet comda.',
      10.305, r.ty, 5.541, 2.719, BODY);
  });
});

// 23 — influencer
build.push(s => {
  txt(s, 'influencer', 1.398, 4.164, 9.354, 1.711, DISPLAY);
  [2.774, 9.530].forEach((x, i) => {
    photo(s, x, 8.156, 4.908, 6.135);
    txt(s, i === 0 ? 'KIM DOMS' : 'MILA MONA', x, 14.708, 4.132, 0.669, BODY, { valign: 'middle' });
    iconSocial(s, 'instagram', x, 15.794, 0.444, 0.444, GRAY);
    txt(s, 'Des diam unsa dniec eirmo di mera cos.', x, 16.942, 2.963, 2.047, BODY);
  });
});

// 24 — collection
build.push(s => {
  txt(s, 'Conm estasmae, notrad seat the os indsa takimta sancem ipsusit drin for themas amet comda.',
    1.398, 2.774, 7.213, 2.036, BODY);
  photo(s, 11.660, 6.081, 5.556, 5.556);
  txt(s, 'collection', 1.398, 8.500, 9.542, 1.711, DISPLAY, { valign: 'middle' });
  photo(s, 0, 11.637, 5.556, 6.944);
  photo(s, 11.660, 12.049, 5.562, 6.944);
  photo(s, 0, 18.993, 5.556, 5.368);
  txt(s, 'Clean and minimal design.', 11.660, 20.656, 3.236, 2.317, HEAD);
});

// 25 — checklist
build.push(s => {
  txt(s, 'Modern clean and minimal design.', 7.516, 4.164, 6.444, 1.604, HEAD);
  blob(s, 'pebble', 11.102, 7.511, 3.776, 3.120, GOLD, 300.96);
  txt(s, 'checklist', 7.516, 7.956, 8.330, 1.711, DISPLAY);
  [['nosotm espo seat the template design by pixasquare for the druamet comda.', 13.098],
   ['cunsda nros seat the os takimta sancem ipsusit drisn for the druamet comda.', 16.298],
   ['pruda seatm the os takimta sancem ipsusit drin for est the druamet comda.', 19.523],
  ].forEach(row => {
    iconCheck(s, 2.774, row[1], 0.889, BLACK);
    txt(s, 'Deot clita kuasd gubergren, ' + row[0], 4.607, row[1] + 0.025, 7.833, 2.036, BODY);
  });
});

// 26 — our goals
build.push(s => {
  txt(s, ['our', 'goals'], 1.398, 5.748, 6.646, 3.161, DISPLAY);
  photo(s, 9.530, 2.774, 4.908, 6.135);
  [['48k', 12.632, 'Donas the clita on kuasd consas dusa, nosotrad set the os takimta susito drin for the druamet comda.'],
   ['7k', 16.076, 'Donas the clita consas dusa, nosotad set hem os takimta susito idam osdra croan drin for comda.'],
   ['131k', 19.548, 'Donas the cliata kuasd consas dusa, template design by pixasquare susito drin for the druamet comda.'],
  ].forEach(row => {
    txt(s, row[0], 2.774, row[0] === '131k' ? 19.651 : row[1], 3.746, 1.711, DISPLAY);
    txt(s, row[2], 7.232, row[1], 7.206, 2.036, BODY);
  });
});

// 27 — mosy design
build.push(s => {
  photo(s, 0, 7.973, 11.888, 15.0);
  txt(s, ['mosy', 'design'], 1.398, 2.774, 8.331, 3.161, DISPLAY);
  blob(s, 'pebble', 7.307, 13.659, 7.971, 6.586, TEAL, 51.22);
  txt(s, 'Modern clean creative design.', 10.152, 13.869, 5.694, 1.604, HEAD);
  txt(s, 'Maorem eso ipsum dolor sito amet, derasm the conseteto sadipscing elitromas. Sedma diam dnic nonumy eirmod tempor invidun.',
    10.152, 16.047, 5.694, 3.414, BODY);
});

// 28 — S.W.O.T.
build.push(s => {
  blob(s, 'wave', 4.169, 2.463, 5.691, 3.463, TEAL, 135.0);
  blob(s, 'pebble', 4.187, 17.775, 4.723, 3.902, GOLD, 345.96);
  [{ k: 'S', ky: 4.155, kx: 3.979, hy: 3.264, hx: 6.507, title: 'Domas deras idam', by: 4.564, copy: 'nosm' },
   { k: 'O', ky: 9.141, kx: 3.882, hy: 8.249, hx: 6.410, title: 'Cora leada yosa', by: 9.550, copy: 'nosm' },
   { k: 'W', ky: 14.188, kx: 3.882, hy: 13.235, hx: 6.410, title: 'Fros cors werasm', by: 14.535, copy: 'notrad' },
   { k: 'T', ky: 19.201, kx: 3.979, hy: 18.248, hx: 6.507, title: 'Osma pera lorne', by: 19.548, copy: 'notrad' },
  ].forEach(r => {
    txt(s, r.k, r.kx, r.ky, 2.5, 1.711, DISPLAY);
    txt(s, r.title, r.hx, r.hy, 6.736, 0.892, HEAD);
    txt(s, 'Deot clita kuasd gubergren, ' + r.copy + ' seat the os takimta sancem ipsusit drin for the druamet comda.',
      r.hx, r.by, 6.597, 2.036, BODY);
  });
});

// 29 — mosy style
build.push(s => {
  blob(s, 'pebble', 1.154, 2.971, 3.687, 3.047, GOLD, 290.56);
  txt(s, 'Modern clean creative design.', 2.943, 3.892, 6.333, 1.604, HEAD);
  txt(s, ['mosy', 'style'], 1.398, 8.668, 5.362, 3.161, DISPLAY, { valign: 'middle' });
  photo(s, 10.679, 8.668, 5.167, 6.458);
  photo(s, 1.376, 15.126, 5.383, 6.458);
  photo(s, 10.679, 15.126, 5.167, 6.458);
});

// 30 — service A/B/C/D
build.push(s => {
  txt(s, 'service', 1.398, 2.774, 7.935, 1.711, DISPLAY);
  [{ k: 'A.', x: 1.398, y: 7.045, tx: 3.926, ty: 7.453, copy: 'nosm seat the os takimta sancem' },
   { k: 'B.', x: 1.398, y: 11.128, tx: 3.926, ty: 11.537, copy: 'nosm seat design by pixasquare' },
   { k: 'C.', x: 6.721, y: 15.233, tx: 9.249, ty: 15.580, copy: 'notrad seat the os takimta sancem' },
   { k: 'D.', x: 6.721, y: 19.084, tx: 9.249, ty: 19.432, copy: 'notrad seat the os takimta sancem' },
  ].forEach(r => {
    txt(s, r.k, r.x, r.y, 2.5, 1.711, DISPLAY);
    txt(s, 'Deot clita kuasd gubergren, ' + r.copy + ' ipsusit drin for the druamet comda.', r.tx, r.ty, 6.597, 2.036, BODY);
  });
  photo(s, 13.346, 7.453, 2.5, 2.5);
  photo(s, 1.398, 15.233, 2.5, 2.5);
  photo(s, 1.426, 19.084, 2.5, 2.5);
});

// 31 — mobile style
build.push(s => {
  txt(s, ['mobile', 'style'], 1.398, 2.774, 7.567, 3.161, DISPLAY);
  phoneMock(s, 1.488, 10.393, 5.311, 10.747, PANEL);
  photo(s, 1.706, 10.597, 4.869, 10.312, { round: true });
  dropCaption(s, 1.706, 10.597, 4.869, 10.312);
  txt(s, 'Notrad seat the osma indsa takimta the sancem ips drin for amet on comda.', 8.965, 5.935, 5.473, 2.036, BODY);
  txt(s, 'new styled app', 8.965, 10.250, 5.541, 0.892, HEAD, { valign: 'middle' });
  [12.800, 16.029].forEach(y => {
    iconCheck(s, 8.965, y, 0.889, INK);
    txt(s, 'takimta sanc epsuso drin for the druamet uscma comda.', 10.521, y - 0.139, 3.986, 2.036, BODY);
  });
});

// 32 — 4 year plan
build.push(s => {
  blob(s, 'wave', 1.407, 2.174, 6.453, 3.926, TEAL, 135.0);
  txt(s, ['4 year', 'plan'], 1.398, 3.193, 5.403, 3.161, DISPLAY, { valign: 'middle' });
  txt(s, ['Do mas ios', 'idam orac.'], 10.994, 4.164, 4.458, 1.604, HEAD, { valign: 'middle' });
  [['2021', 6.443, 7.514], ['2022', 10.994, 7.514], ['2023', 6.443, 15.718], ['2024', 10.994, 15.718]]
    .forEach(c => txt(s, c[0], c[1], c[2], 3.444, 0.892, HEAD, { valign: 'middle' }));
  [[6.443, 8.888], [10.994, 8.888], [6.443, 17.091], [10.994, 17.091]].forEach(p => photo(s, p[0], p[1], 3.444, 3.444));
  [[6.443, 12.733], [10.994, 12.733], [6.443, 20.937], [10.994, 20.937]].forEach(p =>
    txt(s, ['Deso tams odna', 'idnas incm corna', 'con rasves.'], p[0], p[1], 3.444, 2.036, BODY));
});

// 33 — tablet style
build.push(s => {
  txt(s, ['tablet', 'style'], 2.774, 2.774, 6.123, 3.161, DISPLAY);
  photo(s, 2.774, 8.676, 3.504, 3.504);
  tabletMock(s, 10.220, 10.043, 8.367, 11.585, PANEL);
  photo(s, 10.663, 10.371, 7.542, 10.806, { round: true });
  dropCaption(s, 10.663, 10.371, 7.542, 10.806);
  txt(s, 'Modern clean and minimal design.', 2.774, 13.573, 3.236, 3.029, HEAD);
  txt(s, 'Notrad seat the os indsa takimta sancem ipsusit drin for themas amet comda.', 2.774, 18.182, 3.236, 3.403, BODY);
});

// 34 — plan 1/2/3
build.push(s => {
  txt(s, 'plan ', 2.774, 4.058, 5.625, 1.711, DISPLAY);
  txt(s, 'Modern clean and minimal design.', 9.179, 4.164, 6.667, 1.604, HEAD, { valign: 'middle' });
  [['1.', 9.279], ['2.', 13.636], ['3.', 17.890]].forEach((r, i) => {
    txt(s, r[0], 2.774, r[1], 2.5, 1.711, DISPLAY);
    txt(s, 'Nosm seat the os takimta sancem ipsusit drin for the druamet comda.',
      5.274, [9.279, 13.584, 17.890][i], 5.0, 2.036, BODY);
    photo(s, 11.277, [9.279, 13.584, 17.890][i], 4.569, 2.222);
  });
});

// 35 — watch mock-up
build.push(s => {
  blob(s, 'drop', 1.286, 10.268, 5.510, 4.080, GOLD, 142.43, true);
  txt(s, 'Modern clean and minimal design.', 9.102, 4.164, 5.336, 2.317, HEAD, { valign: 'middle' });
  txt(s, 'Deot clita kuasd conm  desn estasmae, notrad seat the os indsa takimta sancem ipsusit drin for amet comda.',
    9.102, 7.655, 5.541, 2.719, BODY);
  watchMock(s, 1.939, 7.399, 3.367, 5.791, PANEL, BLACK);
  photo(s, 2.172, 8.620, 2.808, 3.321, { round: true });
  iconCheck(s, 9.102, 11.687, 0.889, GRAY);
  txt(s, 'takimta sanc epsuso drin for the druamet uscma comda.', 10.658, 11.548, 3.986, 2.036, BODY);
  photo(s, 1.398, 18.059, 3.525, 3.525);
  photo(s, 10.913, 14.613, 3.525, 6.891);
});

// 36 — pricing
build.push(s => {
  txt(s, ['our', 'pricing'], 1.398, 4.164, 6.283, 3.161, DISPLAY);
  txt(s, 'Desma nora.', 11.652, 5.043, 4.194, 0.892, HEAD);
  blob(s, 'drop', 0.399, 10.140, 7.088, 5.248, GOLD, 142.43, true);
  [{ x: 3.688, name: 'Starter Pack', price: '$ 39', copy: 'design by pixasquare sorem' },
   { x: 10.846, name: 'Agency Pack', price: '$ 79', copy: 'nosotrad sea the takimta sorem' },
  ].forEach(c => {
    txt(s, c.name, c.x, 11.181, 5.0, 0.892, HEAD, { align: 'center' });
    txt(s, c.price, c.x, 12.498, 5.0, 1.711, DISPLAY, { align: 'center' });
    txt(s, 'Colita kuasd bergren, ' + c.copy + ' ipsusit drin for the druamet.', c.x, 14.856, 5.0, 2.719, BODY_GRAY, { align: 'center' });
  });
});

// 37 — contact
build.push(s => {
  txt(s, 'contact', 1.398, 2.453, 7.343, 1.711, DISPLAY);
  txt(s, 'Maorem eso psum dolor sito amet, derasco the cons esdcas elitroma erda cons mirda the onas corda.',
    9.175, 5.875, 6.671, 2.047, BODY);
  txt(s, 'www.company.com', 1.398, 8.688, 5.694, 0.681, BODY);
  txt(s, 'mail@company.com', 9.175, 8.688, 5.694, 0.681, BODY);
  txt(s, 'Follow us to check all our new style', 1.398, 10.668, 5.694, 1.604, HEAD);
  photo(s, 4.245, 13.573, 11.601, 9.364);
  blob(s, 'pebble', 1.293, 16.083, 4.723, 3.902, TEAL, 9.46);
  iconSocial(s, 'twitter', 2.270, 17.951, 0.664, 0.541, BLACK);
  iconSocial(s, 'instagram', 3.387, 17.888, 0.667, 0.667, BLACK);
  iconSocial(s, 'youtube', 4.506, 17.988, 0.667, 0.467, BLACK);
});

// 38 — mock-up library (light screens)
build.push(s => {
  txt(s, 'MOCKUPS', 1.688, 1.455, 2.107, 0.601, { fontFace: 'Montserrat Light', fontSize: 27, color: GRAY, lineSpacingMultiple: 0.9 });
  phoneMock(s, 2.774, 3.294, 4.769, 2.357, PANEL, true);
  phoneMock(s, 8.963, 3.294, 2.357, 4.769, PANEL);
  phoneMock(s, 12.853, 3.294, 2.071, 4.179, PANEL);
  watchMock(s, 2.928, 6.965, 1.562, 2.686, PANEL, DEV_EDGE);
  watchMock(s, 5.774, 6.958, 1.562, 2.686, PANEL, 'F5EFEA');
  tabletMock(s, 2.802, 10.905, 6.404, 4.625, PANEL);
  tabletMock(s, 10.398, 9.074, 4.625, 6.404, PANEL);
  laptopMock(s, 2.199, 16.654, 10.045, 5.639, 'FFFFFF');
  phoneMock(s, 12.952, 16.654, 2.071, 4.179, PANEL);
  [[2.874, 3.402, 4.556, 2.140], [9.066, 3.402, 2.152, 4.525], [12.920, 3.337, 1.902, 4.045],
   [3.030, 7.517, 1.291, 1.571], [5.859, 7.505, 1.323, 1.591], [3.040, 11.137, 5.963, 4.198],
   [10.601, 9.281, 4.203, 5.980], [13.050, 16.707, 1.874, 4.045],
  ].forEach(p => photo(s, p[0], p[1], p[2], p[3], { round: true }));
  photo(s, 3.294, 16.862, 7.856, 4.933);
  [[2.874, 3.402, 4.556, 2.140], [9.066, 3.402, 2.152, 4.525], [12.920, 3.337, 1.902, 4.045],
   [3.040, 11.137, 5.963, 4.198], [10.601, 9.281, 4.203, 5.980], [13.050, 16.707, 1.874, 4.045],
  ].forEach(p => dropCaption(s, p[0], p[1], p[2], p[3]));
});

// 39 — device library (dark screens)
build.push(s => {
  txt(s, 'DEVICES', 1.688, 1.455, 1.805, 0.601, { fontFace: 'Montserrat Light', fontSize: 27, color: GRAY, lineSpacingMultiple: 0.9 });
  phoneMock(s, 2.774, 3.252, 4.704, 2.325, DEV_DARK, true);
  phoneMock(s, 8.877, 3.252, 2.378, 4.812, DEV_DARK);
  phoneMock(s, 12.796, 3.252, 2.142, 4.322, DEV_DARK);
  watchMock(s, 2.774, 6.954, 1.539, 2.647, DEV_DARK, DEV_LIGHT);
  watchMock(s, 5.682, 6.954, 1.539, 2.647, DEV_DARK, BLACK);
  tabletMock(s, 2.774, 10.829, 6.392, 4.616, DEV_DARK);
  tabletMock(s, 10.407, 9.053, 4.616, 6.392, DEV_DARK);
  laptopMock(s, 2.004, 16.436, 10.433, 5.857, DEV_DARK);
  phoneMock(s, 12.878, 16.436, 2.060, 4.157, DEV_DARK);
});

// 40-43 — bonus line-icon sheets
build.push(s => {
  txt(s, 'Vector Line Icons', 1.389, 4.174, 7.381, 0.699, { fontFace: 'Lato Bold', fontSize: 34, color: INK, lineSpacingMultiple: 0.9 });
  iconSheet(s, 1.575, [{
    label: 'General', labelX: 1.520, labelY: 6.479, set: 'general',
    rows: [[8.018, 12], [9.044, 12], [10.073, 12], [11.105, 12], [12.140, 12], [13.164, 12],
      [14.197, 12], [15.226, 12], [16.279, 12], [17.283, 12]],
  }]);
});

build.push(s => {
  iconSheet(s, 1.486, [
    { label: 'Time / Location', labelY: 4.139, set: 'place', rows: [[5.617, 12], [6.655, 8]] },
    { label: 'Design', labelY: 7.859, set: 'design', rows: [[9.205, 11]] },
    { label: 'Weather', labelY: 10.465, set: 'weather', rows: [[11.938, 9]] },
    { label: 'Message', labelY: 13.128, set: 'message', rows: [[14.493, 12], [15.455, 3]] },
    { label: 'People', labelY: 16.732, set: 'people', rows: [[18.029, 12], [19.011, 3]] },
  ]);
});

build.push(s => {
  iconSheet(s, 1.482, [
    { label: 'Smileys', labelY: 4.139, set: 'smiley', rows: [[5.578, 11], [6.573, 2]] },
    { label: 'Arrows', labelY: 8.136, set: 'arrows',
      rows: [[9.669, 12], [10.695, 12], [11.737, 12], [12.762, 11], [13.792, 12], [14.835, 11]] },
    { label: 'Shopping / Finance', labelY: 16.950, set: 'shop',
      rows: [[18.446, 12], [19.489, 12], [20.503, 12], [21.525, 2]] },
  ]);
});

build.push(s => {
  iconSheet(s, 1.520, [
    { label: 'Devices', labelX: 1.408, labelY: 4.196, set: 'devices', rows: [[5.542, 12], [6.556, 12], [7.582, 12], [8.526, 9]] },
    { label: 'Social Media', labelY: 10.582, set: 'circle', rows: [[12.057, 12], [13.102, 12], [14.107, 12], [15.052, 3]] },
    { set: 'socialGlyph', filled: true, rows: [[16.988, 12], [18.002, 12], [19.044, 12], [20.022, 3]] },
  ]);
});

/* ---------------------------------------------------------------------- main */
function main() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'MOSY', width: W, height: H });
  pres.layout = 'MOSY';
  pres.author = 'Pixasquare';
  pres.title = 'MOSY';
  build.forEach(fn => {
    const s = pres.addSlide();
    s.background = { color: 'FFFFFF' };
    fn(s);
  });
  return pres.writeFile({ fileName: path.join(__dirname, '0014a4ae-2079-42e1-b580-7967ce4300bf_grok_final.pptx') });
}

main().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
