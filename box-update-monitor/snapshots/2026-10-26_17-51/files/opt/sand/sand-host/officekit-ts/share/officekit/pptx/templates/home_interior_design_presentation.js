/**
 * Recreation of "Gamish - Home Interior Presentation Template" (30 slides, 13.333 x 7.5 in)
 * with pptxgenjs.  Raster images in the source deck are replaced by flat placeholder shapes.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const C = {
  accent: 'B47654', // theme accent1
  accentDk: '87583F', // accent1 shade 75%
  accentLt: 'E1C8BB',
  gray1: 'D8D8D8', // accent2
  gray2: 'BFBFBF', // accent3
  gray3: 'A5A5A5', // accent4
  gray4: '7F7F7F', // accent5
  gray5: '595959', // accent6 / body text
  dark: '404040',
  darker: '262626',
  white: 'FFFFFF',
  paper: 'F2F2F2',
  line: 'D9D9D9',
  muted: '808080'
};
const HEAD = 'Mulish Medium'; // theme major font
const BODY = 'Open Sans'; // theme minor font

/* ------------------------------------------------------------------- copy */
const LOREM = {
  long: 'dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ' +
        'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.',
  mid: 'dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ' +
       'Ut enim ad minim veniam, quis nostrud exercitation',
  short: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
  tiny: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do',
  two: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, sed do',
  card: 'dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt',
  cut: 'dolor sit amet, consectetur adip iscing elit, sed do eiusmod',
  cut2: 'dolor sit amet, consectetur adipi scing elit, sed do eiusmod'
};
/** "Lorem ipsum" in bold followed by regular copy - used all over the deck. */
const lead = rest => [{ text: 'Lorem ipsum ', options: { bold: true } }, { text: rest, options: { bold: false } }];

/* ---------------------------------------------------------------- helpers */
const NOLINE = { type: 'none' };

/** Body copy: 12pt Open Sans, grey, top aligned, 130% line spacing. */
function body(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: BODY, fontSize: 12, color: C.gray5, valign: 'top', lineSpacingMultiple: 1.3
  }, o));
}
/** Display heading: Mulish Medium. */
function head(s, text, o) {
  s.addText(text, Object.assign({
    fontFace: HEAD, fontSize: 44, color: C.gray5, valign: 'top', lineSpacingMultiple: 0.9
  }, o));
}
function rect(s, x, y, w, h, fill, o) {
  s.addShape('rect', Object.assign({ x, y, w, h, fill, line: NOLINE }, o));
}
function rrect(s, x, y, w, h, fill, radius, o) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, fill, rectRadius: radius, line: NOLINE }, o));
}
/** Stand-in for a photo/logo from the source deck. */
function imgBox(s, x, y, w, h, o) {
  o = o || {};
  s.addShape(o.round ? 'roundRect' : (o.oval ? 'ellipse' : 'rect'), {
    x, y, w, h, rectRadius: o.round || 0,
    fill: { color: o.color || C.paper, transparency: o.transparency === undefined ? 60 : o.transparency },
    line: o.line || NOLINE, rotate: o.rotate || 0
  });
  if (o.label) {
    s.addText('[image]', {
      x, y, w, h, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 11, color: o.labelColor || C.gray3
    });
  }
}
/** Thin line-art glyph placeholder standing in for the deck's icon artwork. */
function icon(s, x, y, sz, color) {
  s.addShape('roundRect', { x, y, w: sz, h: sz, fill: NOLINE, line: { color, width: 1.25 }, rectRadius: sz * 0.16 });
  s.addShape('ellipse', { x: x + sz * 0.2, y: y + sz * 0.2, w: sz * 0.26, h: sz * 0.26, fill: NOLINE, line: { color, width: 1 } });
  s.addShape('rect', { x: x + sz * 0.2, y: y + sz * 0.62, w: sz * 0.6, h: sz * 0.14, fill: { color }, line: NOLINE });
}
/** Blend two hex colours; t = 0 -> a, 1 -> b. */
function mix(a, b, t) {
  const ch = i => Math.round(parseInt(a.substr(i, 2), 16) + (parseInt(b.substr(i, 2), 16) - parseInt(a.substr(i, 2), 16)) * t);
  return [ch(0), ch(2), ch(4)].map(v => v.toString(16).padStart(2, '0').toUpperCase()).join('');
}
/** Colour at position t across a list of [pos, hex] stops. */
function stopColor(stops, t) {
  for (let i = 1; i < stops.length; i++) {
    if (t <= stops[i][0] || i === stops.length - 1) {
      const [p0, c0] = stops[i - 1], [p1, c1] = stops[i];
      return mix(c0, c1, p1 === p0 ? 0 : (t - p0) / (p1 - p0));
    }
  }
  return stops[0][1];
}
/** pptxgenjs has no gradient fill: paint a vertical gradient as stacked bands. */
function gradientV(s, x, y, w, h, stops, steps) {
  steps = steps || 48;
  for (let i = 0; i < steps; i++) {
    rect(s, x, y + h * i / steps, w, h / steps + 0.02, { color: stopColor(stops, (i + 0.5) / steps) });
  }
}
/** Left-to-right gradient on a rounded card: nested roundRects sharing the left edge. */
function gradientCardH(s, x, y, w, h, radius, from, to, steps) {
  steps = steps || 32;
  for (let i = steps; i >= 1; i--) {
    rrect(s, x, y, w * i / steps, h, { color: mix(from, to, (i - 1) / (steps - 1)) }, radius);
  }
}
/** Small "v" chevron used as a scroll hint. */
function chevron(s, x, y, w, h, color) {
  s.addShape('custGeom', {
    x, y, w, h, fill: { color }, line: NOLINE,
    points: [{ x: 0, y: 0 }, { x: w / 2, y: h * 0.72 }, { x: w, y: 0 }, { x: w, y: h * 0.28 },
      { x: w / 2, y: h }, { x: 0, y: h * 0.28 }, { close: true }]
  });
}
/** Circle enclosing a right arrow (card "go" affordance). */
function arrowCircle(s, x, y, sz, color) {
  s.addShape('ellipse', { x, y, w: sz, h: sz, fill: NOLINE, line: { color, width: 1.5 } });
  s.addShape('line', { x: x + sz * 0.24, y: y + sz * 0.5, w: sz * 0.42, h: 0, line: { color, width: 1.25 } });
  s.addShape('triangle', { x: x + sz * 0.58, y: y + sz * 0.36, w: sz * 0.28, h: sz * 0.28, fill: { color }, line: NOLINE, rotate: 90 });
}
/** Pill button, e.g. "Read more". */
function button(s, x, y, w, h, fill, color) {
  rrect(s, x, y, w, h, { color: fill }, 0.1);
  s.addText('Read more', { x, y, w, h, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 10.5, italic: true, color });
}
/** Label + percentage above a horizontal progress track. */
function statBar(s, o) {
  const lineY = o.y + (o.lineDy === undefined ? 0.465 : o.lineDy);
  const lw = o.w - 0.19;
  body(s, o.label, { x: o.x, y: o.y, w: 1.392, h: 0.341, fontSize: 12, color: o.labelColor || C.gray5 });
  body(s, [{ text: String(o.pct) }, { text: '%', options: { superscript: true } }],
    { x: o.x + o.w - 1.392, y: o.y, w: 1.392, h: 0.341, align: 'right', color: o.valueColor || C.gray5 });
  s.addShape('line', { x: o.x + 0.09, y: lineY, w: lw, h: 0, line: { color: o.track || C.paper, width: 4 } });
  s.addShape('line', { x: o.x + 0.09, y: lineY, w: lw * o.pct / 100, h: 0, line: { color: o.bar, width: 4 } });
}
/** Ring gauge: pie slice + white centre disc, matching the source deck's charts. */
function ring(s, x, y, sz, pct, color, label, o) {
  o = o || {};
  s.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [pct / 100, 1 - pct / 100] }], {
    x, y, w: sz, h: sz, chartColors: [color, o.rest || C.paper], showLegend: false, showTitle: false,
    dataBorder: { pt: 0, color: 'FFFFFF' }, chartArea: { fill: { color: 'FFFFFF', transparency: 100 } }
  });
  const hole = sz * (o.hole || 0.66);
  s.addShape('ellipse', { x: x + (sz - hole) / 2, y: y + (sz - hole) / 2, w: hole, h: hole, fill: { color: C.white }, line: NOLINE });
  s.addText([{ text: label }, { text: '%', options: { superscript: true } }], {
    x: x + (sz - hole) / 2, y: y + (sz - hole) / 2, w: hole, h: hole, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: o.fontSize || 16, color: C.gray5, margin: 0
  });
}
/** Every slide carries the master's page number behind its artwork. */
function newSlide(pres, num) {
  const s = pres.addSlide();
  s.addText(String(num), {
    x: 0.195, y: 6.966, w: 0.61, h: 0.37, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 16, color: C.muted
  });
  return s;
}

/* =========================================================== slide builders */

// 1 - Cover
function slide01(p) {
  const s = newSlide(p, 1);
  gradientV(s, 0, 0, 13.333, 7.5, [[0, 'C5957B'], [0.5, 'BC8568'], [1, 'B57655']], 60);
  head(s, 'Gamish', { x: 2.866, y: 2.374, w: 7.602, h: 2.423, fontSize: 138, color: C.white, align: 'center', lineSpacingMultiple: 1 });
  body(s, 'Home Interior Presentation Template', {
    x: 3.156, y: 4.567, w: 7.021, h: 0.462, fontSize: 18, color: C.white, align: 'center', paraSpaceAfter: 6
  });
  chevron(s, 6.549, 6.065, 0.236, 0.136, C.gray2);
}

// 2 - About Us
function slide02(p) {
  const s = newSlide(p, 2);
  head(s, 'About Us\nAnd Our Company', { x: 1.115, y: 1.256, w: 6.433, h: 1.582, lineSpacingMultiple: 1 });
  body(s, lead(LOREM.long), { x: 1.115, y: 3.013, w: 10.282, h: 0.604 });
  const cards = [
    { x: 1.214, fill: C.accent, fg: C.white },
    { x: 7.029, fill: C.gray1, fg: C.gray5 }
  ];
  cards.forEach(c => {
    rrect(s, c.x, 4.314, 5.09, 1.887, { color: c.fill }, 0.148);
    icon(s, c.x + 0.4, 4.535, 0.37, c.fg);
    body(s, 'Your text in here', { x: c.x + 0.99, y: 4.45, w: 3.178, h: 0.422, fontSize: 16, bold: true, color: c.fg, paraSpaceAfter: 12 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.',
      { x: c.x + 0.39, y: 5.022, w: 4.185, h: 0.866, color: c.fg });
  });
}

// 3 - Our Main Service (4 columns)
function slide03(p) {
  const s = newSlide(p, 3);
  head(s, 'Our Main Service', { x: 2.695, y: 0.999, w: 7.943, h: 0.841, align: 'center', lineSpacingMultiple: 1 });
  body(s, lead(LOREM.mid), { x: 2.785, y: 1.997, w: 7.752, h: 0.604, align: 'center' });
  const cols = [
    { x: 0.763, icon: 0.899, title: 'Architecture Design', rule: null, color: C.accent },
    { x: 3.833, icon: 3.976, title: 'Interior\nDesign', rule: 3.591, color: 'D2AD98' },
    { x: 6.919, icon: 7.039, title: 'Product\nDesign', rule: 6.661, color: 'E8E8E8' },
    { x: 10.016, icon: 10.123, title: 'Project\nManagement', rule: 9.73, color: 'C9C9C9' }
  ];
  cols.forEach((c, i) => {
    if (c.rule) s.addShape('line', { x: c.rule, y: 3.218, w: 0, h: 4.282, line: { color: c.color, width: 2 } });
    icon(s, c.icon, 3.6, 0.62, i === 0 ? C.accent : C.gray2);
    body(s, c.title, { x: c.x + 0.043, y: 4.492, w: 1.99, h: 0.64, fontSize: 16, bold: true, lineSpacingMultiple: 1 });
    body(s, lead(LOREM.card), { x: c.x, y: i === 0 ? 5.271 : 5.244, w: 2.685, h: 0.866 });
  });
}

// 4 - Our Story
function slide04(p) {
  const s = newSlide(p, 4);
  rect(s, 0, 0, 13.333, 3.75, { color: C.accent });
  head(s, 'Our Story', { x: 6.418, y: 1.264, w: 6.13, h: 0.919, fontSize: 54, color: C.white });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris'),
    { x: 6.441, y: 2.397, w: 6.158, h: 0.866, color: C.white });
  [{ x: 6.56, ic: C.accent }, { x: 9.349, ic: C.gray3 }].forEach(c => {
    icon(s, c.x, 4.248, 0.45, c.ic);
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', { x: c.x - 0.119, y: 4.946, w: 2.397, h: 0.866 });
    button(s, c.x, 5.978, 1.156, 0.4, C.accent, C.white);
  });
}

// 5 - Our Branding (3 cards + rotated title)
function slide05(p) {
  const s = newSlide(p, 5);
  const cards = [
    { x: 1.88, fill: C.accent, fg: C.white, arrow: C.white },
    { x: 5.54, fill: C.gray1, fg: C.gray5, arrow: C.gray5 },
    { x: 9.21, fill: C.gray3, fg: C.white, arrow: C.white }
  ];
  cards.forEach(c => {
    rrect(s, c.x, 1.444, 3.377, 4.944, { color: c.fill }, 0.166, { line: { color: C.line, width: 1 } });
    imgBox(s, c.x + 0.249, 1.667, 2.879, 2.083, { transparency: 100 });
    body(s, 'Your text in here', { x: c.x + 0.249, y: 4.29, w: 2.879, h: 0.422, fontSize: 16, bold: true, color: c.fg, paraSpaceAfter: 12 });
    body(s, LOREM.short, { x: c.x + 0.249, y: 4.762, w: 2.879, h: 0.866, color: c.fg });
    arrowCircle(s, c.x + 2.857, 5.889, 0.272, c.arrow);
  });
  head(s, 'Our Branding', { x: -1.307, y: 3.271, w: 4.712, h: 0.707, fontSize: 40, align: 'center', rotate: 270 });
}

// 6 - Latest Property News
function slide06(p) {
  const s = newSlide(p, 6);
  rrect(s, 0.82, 0.82, 4.115, 5.917, { color: C.accent }, 0.24);
  icon(s, 1.33, 1.355, 0.579, C.white);
  head(s, 'Latest Property News', { x: 1.194, y: 2.088, w: 3.338, h: 2.322, color: C.white, lineSpacingMultiple: 1.15 });
  [['Project 1', 0.977, 1.421], ['Project 2', 3.145, 3.497], ['Project 3', 5.271, 5.641]].forEach(([t, y1, y2]) => {
    body(s, t, { x: 5.421, y: y1, w: 3.222, h: 0.374, fontSize: 18, lineSpacingMultiple: 0.9 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod Lorem ipsum dolor sit amet, consectetur',
      { x: 5.42, y: y2, w: 3.541, h: 0.866 });
  });
}

// 7 - Our Main Service (3 cards over a photo panel)
function slide07(p) {
  const s = newSlide(p, 7);
  imgBox(s, 0, 0, 5.54, 7.5, { transparency: 70 });
  head(s, 'Our Main Service', { x: 5.919, y: 0.857, w: 5.54, h: 0.909, fontSize: 48, color: C.dark, lineSpacingMultiple: 1 });
  const cards = [
    { x: 2.744, fill: C.accent, fg: C.white, btn: C.accentLt, radius: 0.219 },
    { x: 6.113, fill: C.gray1, fg: C.gray5, btn: 'EFEFEF', radius: 0.183 },
    { x: 9.482, fill: C.gray3, fg: C.white, btn: 'DBDBDB', radius: 0.183 }
  ];
  cards.forEach((c, i) => {
    rrect(s, c.x, 2.602, 3.096, 4.02, { color: c.fill }, c.radius);
    icon(s, c.x + 0.4, 3.24, 0.47, c.fg);
    body(s, 'Your text in here', { x: c.x + 1.03, y: 3.14, w: 1.681, h: 0.772, fontSize: 16, bold: true, color: c.fg, paraSpaceAfter: 12 });
    body(s, lead(i === 2 ? LOREM.cut2 : LOREM.cut), { x: c.x + 0.29, y: i === 0 ? 4.186 : 4.167, w: 2.37, h: 0.866, color: c.fg });
    button(s, c.x + 0.935, 5.728, 1.156, 0.4, c.btn, C.gray5);
  });
}

// 8 - Speaker profile with stat bars
function slide08(p) {
  const s = newSlide(p, 8);
  rect(s, 0, 0.75, 13.333, 6.0, { color: C.accent });
  imgBox(s, 1.054, 1.377, 1.62, 1.62, { color: C.paper, transparency: 60, label: true, labelColor: C.white });
  imgBox(s, 5.188, 0.958, 3.75, 5.583, { transparency: 100, label: true, labelColor: 'C9A18A' });
  s.addText([
    { text: 'Hello, my name is ', options: { fontSize: 28 } },
    { text: '\n', options: { fontSize: 28 } },
    { text: 'Joseph Lewis', options: { fontSize: 40 } }
  ], { x: 0.937, y: 3.349, w: 4.976, h: 1.131, fontFace: HEAD, color: C.white, valign: 'bottom', lineSpacingMultiple: 0.9 });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et'),
    { x: 0.937, y: 4.732, w: 3.621, h: 0.866, color: C.white });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore, '),
    { x: 9.477, y: 2.131, w: 3.189, h: 0.866, color: C.white });
  [['Stats One', 80, '87583F', 4.02], ['Stats Two', 76, 'A2A2A2', 4.725], ['Stats Three', 83, '7C7C7C', 5.429]]
    .forEach(([label, pct, bar, y]) => statBar(s, {
      x: 9.48, y, w: 2.792, label, pct, bar, labelColor: C.white, valueColor: C.white, lineDy: 0.465
    }));
}

// 9 - Meet Our Team (roster)
function slide09(p) {
  const s = newSlide(p, 9);
  rect(s, 0, 0, 4.8, 7.5, { color: C.accent });
  s.addText('Meet Our Team', { x: 0.792, y: 2.332, w: 3.125, h: 1.434, fontFace: HEAD, fontSize: 44, color: C.white, valign: 'bottom', lineSpacingMultiple: 0.9 });
  rect(s, 1.076, 3.792, 0.9, 0.04, { color: C.accent });
  body(s, lead(LOREM.short.replace('Lorem ipsum ', '')), { x: 0.861, y: 4.024, w: 3.473, h: 0.604, color: C.white });
  [['Derrick Xavier', 1.093], ['Helena Anne', 3.08], ['Isaiah Matthews', 5.067]].forEach(([name, y]) => {
    imgBox(s, 5.708, y + 0.074, 1.192, 1.192, { color: C.paper, transparency: 60, label: true });
    body(s, name, { x: 7.271, y, w: 5.104, h: 0.422, fontSize: 16, bold: true, valign: 'bottom' });
    body(s, 'Position Name', { x: 7.271, y: y + 0.345, w: 5.104, h: 0.318, fontSize: 11 });
    body(s, LOREM.short, { x: 7.271, y: y + 0.738, w: 4.403, h: 0.604 });
  });
}

// 10 - Meet Our Team (cards)
function slide10(p) {
  const s = newSlide(p, 10);
  const cards = [
    { x: 2.566, fill: C.accent, fg: C.white },
    { x: 5.366, fill: C.gray1, fg: C.gray5 },
    { x: 8.166, fill: C.gray3, fg: C.white }
  ];
  cards.forEach(c => {
    rrect(s, c.x, 2.654, 2.629, 3.522, { color: c.fill }, 0.112);
    imgBox(s, c.x + 0.33, 2.942, 1.969, 1.413, { transparency: 100 });
    body(s, 'Your Name Here', { x: c.x + 0.33, y: 5.025, w: 1.969, h: 0.379, fontSize: 14, color: c.fg, align: 'center', paraSpaceAfter: 12 });
    body(s, 'Job description', { x: c.x + 0.33, y: 5.353, w: 1.969, h: 0.339, fontSize: 12, italic: true, color: c.fg, align: 'center', paraSpaceAfter: 12 });
    chevron(s, c.x + 1.197, 5.91, 0.236, 0.136, C.white);
  });
  head(s, 'Meet Our Team', { x: 3.5, y: 0.769, w: 6.333, h: 0.782, align: 'center' });
  body(s, LOREM.short, { x: 3.514, y: 1.677, w: 6.333, h: 0.604, align: 'center' });
  ['01', '02', '03'].forEach((t, i) => body(s, t, {
    x: 5.918 + i * 0.552, y: 6.542, w: 0.393, h: 0.303, fontSize: 12, color: i === 1 ? C.accent : 'A6A6A6', lineSpacingMultiple: 1
  }));
}

// 11 - Notebook Mockups
function slide11(p) {
  const s = newSlide(p, 11);
  // laptop mockup: light bezel, dark screen, base bar
  imgBox(s, 3.93, 2.85, 5.48, 3.65, { color: 'CFCFCF', transparency: 0, round: 0.09 });
  imgBox(s, 4.16, 3.02, 5.02, 3.4, { color: '141414', transparency: 0, label: true, labelColor: '8E8E8E' });
  imgBox(s, 3.3, 6.5, 6.75, 0.18, { color: 'C4C4C4', transparency: 0, round: 0.09 });
  s.addText('Notebook Mockups', { x: 3.724, y: 0.912, w: 5.885, h: 0.707, fontFace: HEAD, fontSize: 40, color: C.gray5, valign: 'bottom', align: 'center', lineSpacingMultiple: 0.9 });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, sed do'),
    { x: 3.441, y: 1.829, w: 6.466, h: 0.604, align: 'center' });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed', { x: 0.935, y: 3.249, w: 2.506, h: 0.866 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed', { x: 9.962, y: 3.238, w: 2.457, h: 0.866 });
  statBar(s, { x: 0.92, y: 4.396, w: 2.489, label: 'Stats One', pct: 85, bar: C.accent, valueColor: C.accent, lineDy: 0.581 });
  statBar(s, { x: 0.92, y: 5.1, w: 2.489, label: 'Stats Two', pct: 70, bar: C.gray2, valueColor: C.gray2, lineDy: 0.608 });
  [4.514, 4.816, 5.108].forEach((y, i) => {
    s.addShape('ellipse', { x: 10.072 + i * 0.011, y, w: 0.146, h: 0.146, fill: { color: C.accent }, line: NOLINE });
    body(s, 'Lorem ipsum dolor sit amet, ', { x: 10.218 + i * 0.011, y: y - 0.125, w: 2.448, h: 0.301, fontSize: 10 });
  });
}

// 12 - Smartphone Mockups
function slide12(p) {
  const s = newSlide(p, 12);
  [[7.1, 354], [9.72, 6]].forEach(([x, rot]) => {
    imgBox(s, x, 1.0, 2.5, 5.5, { color: '25262A', transparency: 0, round: 0.32, rotate: rot });
    imgBox(s, x + 0.11, 1.12, 2.28, 5.26, { color: 'F5F5F5', transparency: 0, round: 0.26, rotate: rot, label: true, labelColor: '9A9A9A' });
  });
  s.addText('Smartphone Mockups', { x: 0.893, y: 0.928, w: 4.796, h: 1.434, fontFace: HEAD, fontSize: 44, color: C.gray5, valign: 'bottom', lineSpacingMultiple: 0.9 });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, sed do'),
    { x: 0.897, y: 2.736, w: 4.126, h: 0.866 });
  statBar(s, { x: 0.887, y: 4.106, w: 4.796, label: 'Stats One', pct: 90, bar: C.accent, labelColor: C.accent, valueColor: C.accent });
  statBar(s, { x: 0.887, y: 4.81, w: 4.796, label: 'Stats Two', pct: 86, bar: C.gray1 });
  statBar(s, { x: 0.887, y: 5.514, w: 4.796, label: 'Stats Three', pct: 91, bar: C.gray2 });
}

// 13 - Smartwatch Mockups
function slide13(p) {
  const s = newSlide(p, 13);
  // watch mockup: looping dark strap (band + wrist opening) with a light square face on top
  imgBox(s, 8.75, 1.17, 3.5, 5.12, { color: '232323', transparency: 0, round: 1.7 });
  imgBox(s, 9.72, 2.3, 1.45, 3.35, { color: 'FFFFFF', transparency: 0, round: 0.7 });
  imgBox(s, 7.76, 1.85, 2.9, 2.95, { color: '4A4A4A', transparency: 0, round: 0.3, rotate: 9 });
  imgBox(s, 7.96, 2.05, 2.5, 2.55, { color: 'D5D5D5', transparency: 0, round: 0.08, rotate: 9, label: true, labelColor: '6E6E6E' });
  s.addShape('roundRect', { x: 10.44, y: 2.6, w: 0.2, h: 0.36, fill: { color: '5A5A5A' }, line: NOLINE, rectRadius: 0.06, rotate: 9 });
  s.addText('Smartwatch Mockups', { x: 0.893, y: 0.928, w: 4.796, h: 1.434, fontFace: HEAD, fontSize: 44, color: C.gray5, valign: 'bottom', lineSpacingMultiple: 0.9 });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, sed do'),
    { x: 0.999, y: 2.668, w: 3.76, h: 0.866 });
  [{ x: 1.009, pct: 50, color: C.accent, title: 'Statistic One', tc: C.accent },
    { x: 4.353, pct: 70, color: C.gray1, title: 'Statistic Two', tc: C.gray5 }].forEach(g => {
    ring(s, g.x, 4.029, 1.096, g.pct, g.color, String(g.pct), { rest: 'F2F2F2', hole: 0.68 });
    body(s, g.title, { x: g.x, y: 5.157, w: 2.381, h: 0.378, fontSize: 14, bold: true, color: g.tc });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do', { x: g.x, y: g.x < 2 ? 5.535 : 5.528, w: 2.485, h: 0.866 });
  });
}

// 14 - Section Break
function slide14(p) {
  const s = newSlide(p, 14);
  rect(s, 0, 0, 13.333, 7.5, { color: C.accent });
  gradientCardH(s, 1.0, 1.448, 11.365, 4.983, 0.2, 'FFFFFF', 'EBDBD3');
  head(s, 'Section Break', { x: 1.625, y: 2.24, w: 10.083, h: 1.582, fontSize: 88, align: 'center', lineSpacingMultiple: 1 });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, sed do'),
    { x: 3.965, y: 3.966, w: 5.403, h: 0.604, align: 'center' });
  ['00', '01', '20', '00'].forEach((t, i) => {
    const x = 4.922 + i * 0.966;
    s.addShape('roundRect', { x, y: 5.427, w: 0.624, h: 0.624, fill: { color: C.white }, line: { color: C.gray2, width: 1 }, rectRadius: 0.1 });
    s.addText(t, { x, y: 5.427, w: 0.624, h: 0.624, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 20, color: C.accent });
    if (i < 3) s.addText(':', { x: x + 0.633, y: 5.44, w: 0.324, h: 0.6, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 28, color: C.accent });
  });
}

// 15 - Single Timeline Slide (U-shaped track)
function slide15(p) {
  const s = newSlide(p, 15);
  head(s, 'Single Timeline Slide', { x: 3.219, y: 0.911, w: 6.896, h: 0.767, align: 'center', valign: 'bottom' });
  // the track: top rail, right-hand U bend, bottom rail
  s.addShape('line', { x: 2.032, y: 3.685, w: 7.313, h: 0, line: { color: '5A3B2A', width: 3 } });
  s.addShape('line', { x: 2.032, y: 5.029, w: 7.313, h: 0, line: { color: '5A3B2A', width: 3 } });
  s.addShape('arc', { x: 8.673, y: 3.685, w: 1.344, h: 1.344, line: { color: '5A3B2A', width: 3 }, angleRange: [270, 90] });
  const top = [
    { x: 1.042, dot: 2.01, date: '2 July 2023', color: C.accent, dotColor: C.accent },
    { x: 4.121, dot: 5.089, date: '12 July 2023', color: '747373', dotColor: C.gray1 },
    { x: 7.2, dot: 8.168, date: '17 July 2023', color: '747373', dotColor: C.gray2 }
  ];
  const bottom = [
    { x: 1.042, dot: 2.01, date: '20 September 2023', color: C.accent, dotColor: C.accent },
    { x: 4.121, dot: 5.089, date: '5 September 2023', color: C.gray5, dotColor: C.gray5 },
    { x: 7.2, dot: 8.168, date: '17 August 2023', color: C.gray4, dotColor: C.gray4 }
  ];
  top.forEach((t, i) => {
    s.addText(t.date, { x: t.x, y: 2.179, w: 2.027, h: 0.383, align: 'center', fontFace: 'Lexend Deca', fontSize: 14, color: t.color, lineSpacingMultiple: 1.3 });
    body(s, LOREM.tiny, { x: [0.815, 3.917, 6.951][i], y: [2.569, 2.559, 2.562][i], w: 2.434, h: 0.866, align: 'center' });
    s.addShape('ellipse', { x: t.dot, y: 3.641, w: 0.09, h: 0.09, fill: { color: t.dotColor }, line: NOLINE });
  });
  bottom.forEach((t, i) => {
    s.addText(t.date, { x: t.x, y: 5.289, w: 2.027, h: 0.383, align: 'center', fontFace: 'Lexend Deca', fontSize: 14, color: t.color, lineSpacingMultiple: 1.3 });
    body(s, LOREM.tiny, { x: [0.833, 3.94, 6.974][i], y: [5.664, 5.655, 5.658][i], w: 2.434, h: 0.866, align: 'center' });
    s.addShape('ellipse', { x: t.dot, y: 4.982, w: 0.09, h: 0.09, fill: { color: t.dotColor }, line: NOLINE });
  });
  s.addShape('ellipse', { x: 9.972, y: 4.314, w: 0.09, h: 0.09, fill: { color: C.gray3 }, line: NOLINE });
  s.addText('2 August 2023', { x: 10.333, y: 3.906, w: 1.979, h: 0.383, fontFace: 'Lexend Deca', fontSize: 14, color: '747373', lineSpacingMultiple: 1.3 });
  body(s, LOREM.tiny, { x: 10.319, y: 4.297, w: 2.434, h: 0.866 });
}

// 16 - Agenda Of This Month (calendar)
function slide16(p) {
  const s = newSlide(p, 16);
  rect(s, 0, 0, 6.709, 7.5, { color: C.accent });
  s.addText('Agenda Of This Month', { x: 0.707, y: 0.971, w: 2.758, h: 2.692, fontFace: HEAD, fontSize: 44, color: C.white, valign: 'top', lineSpacingMultiple: 1.2 });
  icon(s, 0.838, 4.614, 0.534, C.white);
  s.addText('Architecture Design', { x: 0.73, y: 5.374, w: 2.672, h: 0.37, fontFace: 'Lexend Deca', fontSize: 16, color: C.white });
  body(s, LOREM.tiny, { x: 0.712, y: 5.814, w: 2.485, h: 0.866, color: C.white });
  // month / year steppers
  s.addText('February', { x: 4.584, y: 0.771, w: 1.606, h: 0.404, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 18, color: C.white });
  s.addText('2023', { x: 10.44, y: 0.789, w: 1.961, h: 0.404, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 18, color: C.darker });
  [[4.299, 0.83, C.white, 'leftArrow'], [6.19, 0.83, C.white, 'rightArrow'],
    [10.661, 0.831, C.accent, 'leftArrow'], [11.955, 0.831, C.accent, 'rightArrow']].forEach(([x, y, col, sh]) => {
    s.addShape('rect', { x, y, w: 0.216, h: 0.216, fill: NOLINE, line: { color: col, width: 0.75 } });
    s.addShape('triangle', { x: x + 0.062, y: y + 0.056, w: 0.1, h: 0.1, fill: { color: col }, line: NOLINE, rotate: sh === 'leftArrow' ? 270 : 90 });
  });
  // grid
  rect(s, 3.912, 1.297, 8.764, 3.915, { color: 'C9C9C9' });
  rect(s, 3.912, 1.297, 8.764, 0.449, { color: C.gray4 });
  const DOW = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
  const colX = [4.588, 5.684, 6.811, 7.939, 9.066, 10.193, 11.32];
  DOW.forEach((d, i) => s.addText(d, { x: colX[i], y: 1.358, w: 0.711, h: 0.404, align: 'center', fontFace: BODY, fontSize: 18, color: C.white }));
  const weeks = [
    ['27', '28', '29', '30', '31', '1', '2'],
    ['3', '4', '5', '6', '7', '8', '9'],
    ['10', '11', '12', '13', '14', '15', '16'],
    ['17', '18', '19', '20', '21', '22', '23'],
    ['24', '25', '26', '27', '28', '1', '2']
  ];
  const dim = { '0,0': 1, '0,1': 1, '0,2': 1, '0,3': 1, '0,4': 1, '4,5': 1, '4,6': 1 };
  const marked = { '1,5': 1, '2,1': 1, '3,3': 1 }; // 8th, 11th, 20th
  const numX = [4.589, 5.685, 6.949, 8.045, 9.141, 10.237, 11.291];
  const rowY = [2.259, 2.829, 3.421, 4.012, 4.603];
  weeks.forEach((week, r) => week.forEach((d, c) => {
    if (marked[r + ',' + c]) rect(s, numX[c] + 0.184, rowY[r] - 0.02, 0.342, 0.342, { color: C.accent });
    s.addText(d, { x: numX[c], y: rowY[r], w: 0.711, h: 0.303, align: 'center', fontFace: BODY, fontSize: 12, color: dim[r + ',' + c] ? C.muted : C.white });
  }));
  // agenda captions
  [['Agenda #1', 3.807, C.white, 3.817], ['Agenda #2', 6.904, C.gray5, 6.892], ['Agenda #3', 9.8, C.gray5, 9.789]]
    .forEach(([t, x, col, bx], i) => {
      s.addText(t, { x, y: 5.418, w: 2.687, h: 0.406, fontFace: BODY, fontSize: 16, bold: true, color: col, lineSpacingMultiple: 1.2 });
      body(s, LOREM.tiny, { x: bx, y: [5.792, 5.788, 5.809][i], w: 2.485, h: 0.866, color: col });
    });
}

// 17 - Pictures Gallery Slide (3x3 grid of placeholders)
function slide17(p) {
  const s = newSlide(p, 17);
  rrect(s, 0.667, 1.542, 12.0, 4.417, { color: C.accent }, 0.15);
  const cells = [[3.846, 0.968], [6.833, 0.968], [9.82, 0.968], [6.834, 2.87], [9.821, 2.87],
    [3.846, 4.772], [6.833, 4.772], [9.82, 4.772]];
  cells.forEach(([x, y]) => imgBox(s, x, y, 2.722, 1.76, { color: C.paper, transparency: 60, label: true, labelColor: '9A8377' }));
  s.addText('Pictures Gallery Slide', { x: 0.981, y: 2.146, w: 2.771, h: 1.919, fontFace: HEAD, fontSize: 40, color: C.white, valign: 'bottom', lineSpacingMultiple: 0.9 });
  rect(s, 1.217, 4.065, 0.9, 0.04, { color: C.accent });
  body(s, LOREM.tiny, { x: 1.017, y: 4.236, w: 2.485, h: 0.866, color: C.white });
}

// 18 - Gallery Slide (four framed photos)
function slide18(p) {
  const s = newSlide(p, 18);
  [[1.13, 4.655, 3.306, 2.095], [8.403, 4.655, 3.306, 2.095],
    [4.815, 0.75, 3.201, 3.605], [8.452, 0.75, 3.201, 3.605]]
    .forEach(([x, y, w, h]) => imgBox(s, x, y, w, h, { transparency: 100 }));
  head(s, 'Gallery Slide', { x: 0.91, y: 1.044, w: 3.153, h: 1.571, fontSize: 48 });
  icon(s, 4.984, 4.891, 0.5, C.accent);
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, '), { x: 0.983, y: 2.677, w: 3.256, h: 0.866 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, ', { x: 4.87, y: 5.554, w: 3.034, h: 0.866 });
  ['01', '02', '03'].forEach((t, i) => body(s, t, {
    x: 12.024, y: 3.171 + i * 0.594, w: 0.393, h: 0.303, fontSize: 12, color: i === 1 ? C.accent : C.muted, lineSpacingMultiple: 1
  }));
}

// 19 - Gallery Slide (banner)
function slide19(p) {
  const s = newSlide(p, 19);
  rect(s, 0, 2.683, 13.333, 4.817, { color: C.accent });
  [[0.704, 3.257, 3.211, 2.892], [4.014, 3.257, 4.271, 3.493], [8.396, 3.257, 4.271, 3.493]]
    .forEach(([x, y, w, h]) => imgBox(s, x, y, w, h, { transparency: 100 }));
  head(s, 'Gallery Slide', { x: 4.53, y: 0.781, w: 4.307, h: 0.767, align: 'center' });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, sed do'),
    { x: 3.399, y: 1.737, w: 6.57, h: 0.604, align: 'center' });
  ['01', '02', '03'].forEach((t, i) => body(s, t, {
    x: 1.56 + i * 0.552, y: 6.37, w: 0.393, h: 0.303, fontSize: 12, color: i === 1 ? C.white : C.line, lineSpacingMultiple: 1
  }));
}

// 20 - Agenda Details (table + doughnut)
function slide20(p) {
  const s = newSlide(p, 20);
  rect(s, 6.337, 1.045, 6.438, 5.583, { color: C.white });
  const header = ['Title Text Here', 'Data 1', 'Data 2', 'Data 3', 'Data 4', 'Data 5'];
  const rows = [header.map((t, i) => ({
    text: t, options: { fill: { color: C.accent }, color: C.white, align: i === 0 ? 'left' : 'center' }
  }))];
  for (let r = 0; r < 9; r++) {
    const shade = r % 2 === 0 ? { color: C.white } : { color: 'F2F2F2' };
    const row = [{ text: 'Content Here', options: { fill: shade, color: C.gray5 } }];
    for (let c = 1; c < 6; c++) row.push({ text: '', options: { fill: shade } });
    rows.push(row);
  }
  s.addTable(rows, {
    x: 6.481, y: 1.174, w: 6.15, colW: [1.459, 0.938, 0.938, 0.938, 0.938, 0.938], rowH: 0.53,
    fontFace: BODY, fontSize: 12, color: C.gray5, valign: 'middle', margin: [0, 0.08, 0, 0.08],
    border: [{ type: 'none' }, { type: 'solid', color: 'E6E6E6', pt: 0.75 }, { type: 'none' }, { type: 'solid', color: 'E6E6E6', pt: 0.75 }]
  });
  // tick marks overlaid on the data columns
  const ticks = [[0, 0, C.accent], [0, 2, C.gray2], [1, 0, C.accent], [1, 1, C.gray1], [1, 4, C.gray4],
    [2, 4, C.gray4], [2, 2, C.gray2], [2, 3, C.gray3], [3, 1, C.gray1], [3, 3, C.gray3],
    [4, 1, C.gray1], [4, 4, C.gray4], [5, 0, C.accent], [5, 4, C.gray4], [5, 3, C.gray3],
    [6, 1, C.gray1], [6, 4, C.gray4], [6, 2, C.gray2], [7, 1, C.gray1], [7, 2, C.gray2],
    [8, 0, C.accent], [8, 4, C.gray4], [8, 3, C.gray3]];
  ticks.forEach(([r, c, color]) => {
    const x = 8.303 + c * 0.938, y = 1.864 + r * 0.531;
    s.addShape('ellipse', { x, y, w: 0.216, h: 0.216, fill: NOLINE, line: { color, width: 1 } });
    s.addText('\u2713', { x, y, w: 0.216, h: 0.216, align: 'center', valign: 'middle', fontFace: BODY, fontSize: 9, color, margin: 0 });
  });
  s.addText('Agenda Details', { x: 0.931, y: 1.452, w: 5.137, h: 0.828, fontFace: HEAD, fontSize: 48, color: C.dark, valign: 'bottom', lineSpacingMultiple: 0.9 });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor , consectetur adipiscing elit, '),
    { x: 0.931, y: 2.443, w: 5.047, h: 0.604 });
  s.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr', '5th Qtr'], values: [2, 4, 2, 2, 8] }], {
    x: 1.023, y: 4.005, w: 1.306, h: 1.35, holeSize: 70, showLegend: false, showTitle: false,
    chartColors: [C.accent, C.gray1, C.gray2, C.gray3, C.gray4], dataBorder: { pt: 0, color: 'FFFFFF' }
  });
  icon(s, 1.492, 4.535, 0.368, C.dark);
  body(s, 'Feature Details', { x: 2.593, y: 4.056, w: 2.76, h: 0.381, fontSize: 14, bold: true });
  body(s, LOREM.short, { x: 2.582, y: 4.431, w: 2.637, h: 0.866 });
}

// 21 - Office Chart (clustered bar)
function slide21(p) {
  const s = newSlide(p, 21);
  rect(s, 5.981, 0, 7.352, 7.5, { color: C.accent });
  s.addShape('roundRect', { x: 6.311, y: 1.577, w: 6.105, h: 4.679, fill: { color: C.white }, line: { color: C.gray2, width: 1 }, rectRadius: 0.11 });
  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] }
  ], {
    x: 6.699, y: 1.954, w: 5.331, h: 3.926, barDir: 'bar', barGrouping: 'clustered',
    chartColors: [C.gray2, C.gray1, C.accent], barGapWidthPct: 326, barOverlapPct: -58,
    showLegend: false, showTitle: false, valAxisMaxVal: 5, valAxisMajorUnit: 1,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 12, catAxisLabelColor: C.gray5, valAxisLabelColor: C.gray5,
    catAxisLabelFontFace: BODY, valAxisLabelFontFace: BODY,
    valGridLine: { color: 'D9D9D9', size: 0.75 }, catAxisLineColor: 'D9D9D9'
  });
  head(s, 'Office Chart', { x: 0.909, y: 1.837, w: 3.731, h: 0.767 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor , consectetur adipiscing elit, sed do eiusmod tempor',
    { x: 0.941, y: 2.823, w: 4.654, h: 0.866 });
  button(s, 1.031, 4.165, 1.156, 0.4, C.accent, C.white);
  s.addText([{ text: '432', options: { fontSize: 18 } }, { text: 'k', options: { fontSize: 14 } }],
    { x: 0.927, y: 5.132, w: 2.609, h: 0.459, fontFace: BODY, color: C.gray5, valign: 'top', lineSpacingMultiple: 1.3 });
  s.addText([{ text: '$ ', options: { fontSize: 14 } }, { text: '321', options: { fontSize: 18 } }, { text: '.99', options: { fontSize: 14 } }],
    { x: 3.443, y: 5.132, w: 2.021, h: 0.459, fontFace: BODY, color: C.gray5, valign: 'top', lineSpacingMultiple: 1.3 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur', { x: 0.909, y: 5.553, w: 2.11, h: 0.626 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur', { x: 3.443, y: 5.553, w: 2.11, h: 0.626 });
}

// 22 - Infographic Slide (column chart)
function slide22(p) {
  const s = newSlide(p, 22);
  rect(s, 0, 0, 5.352, 7.5, { color: C.accent });
  const cats = ['5/1/2023', '6/1/2023', '7/1/2023'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [55, 32, 40] },
    { name: 'Series 12', labels: cats, values: [32, 30, 22] },
    { name: 'Series 13', labels: cats, values: [15, 65, 60] },
    { name: 'Series 2', labels: cats, values: [12, 20, 26] }
  ], {
    x: 5.796, y: 2.973, w: 6.739, h: 3.749, barDir: 'col', barGrouping: 'clustered',
    chartColors: [C.accent, C.gray1, C.gray2, C.gray3], barGapWidthPct: 80, barOverlapPct: -33,
    showLegend: false, showTitle: false, showValue: false,
    catAxisLabelFontSize: 10.5, valAxisLabelFontSize: 10.5, catAxisLabelColor: C.gray5, valAxisLabelColor: C.gray5,
    catAxisLabelFontFace: BODY, valAxisLabelFontFace: BODY,
    valGridLine: { color: '404040', size: 0.75 }, catAxisLineColor: C.gray5
  });
  head(s, 'Infographic Slide', { x: 5.759, y: 0.857, w: 5.721, h: 0.841, lineSpacingMultiple: 1 });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor , consectetur adipiscing elit, sed do eiusmod tempor',
    { x: 5.759, y: 1.742, w: 6.204, h: 0.604 });
  [['88%', 0.964], ['70%', 2.918], ['65%', 4.919]].forEach(([pct, y]) => {
    s.addText(pct, { x: 1.213, y, w: 1.827, h: 0.626, fontFace: BODY, fontSize: 28, color: C.white, valign: 'top', lineSpacingMultiple: 1.2 });
    s.addText('Your Text Here', { x: 1.213, y: y + 0.566, w: 2.548, h: 0.401, fontFace: BODY, fontSize: 16, bold: true, color: C.white, valign: 'top', lineSpacingMultiple: 1.2 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor , ', { x: 1.213, y: y + 0.966, w: 3.62, h: 0.604, color: C.white });
  });
}

// 23 - Our Price List
function slide23(p) {
  const s = newSlide(p, 23);
  const cards = [
    { x: 1.353, px: 1.449, fill: C.accent, fg: C.white, price: '110', cents: '.00', name: 'Product 01' },
    { x: 5.317, px: 5.41, fill: C.gray1, fg: C.gray5, price: '125', cents: '.00', name: 'Product 02' },
    { x: 9.252, px: 9.353, fill: C.gray3, fg: C.white, price: '135', cents: '.50', name: 'Product 03' }
  ];
  cards.forEach((c, i) => {
    const yTop = i === 0 ? 2.197 : 2.195;
    rrect(s, c.x, yTop, 2.735, 4.357, { color: c.fill }, 0.178);
    imgBox(s, c.px, i === 2 ? 2.401 : 2.375, 2.531, 2.0, { color: C.paper, transparency: 60, round: 0.15, label: true, labelColor: '6E6E6E' });
    s.addText([
      { text: '$', options: { fontSize: 36, superscript: true } },
      { text: c.price, options: { fontSize: 36 } },
      { text: c.cents, options: { fontSize: 20 } }
    ], { x: c.x + 0.081, y: [4.363, 4.389, 4.373][i], w: 2.531, h: 0.822, align: 'center', valign: 'bottom', fontFace: BODY, color: c.fg, lineSpacingMultiple: 1.3 });
    s.addText(c.name, { x: c.x + 0.113, y: [5.223, 5.218, 5.233][i], w: 2.531, h: 0.378, align: 'center', valign: 'bottom', fontFace: BODY, fontSize: 14, color: c.fg, lineSpacingMultiple: 1.3 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur', { x: c.x + 0.015, y: [5.639, 5.638, 5.633][i], w: 2.646, h: 0.604, align: 'center', color: c.fg });
  });
  head(s, 'Our Price List', { x: 3.219, y: 0.869, w: 6.896, h: 0.767, align: 'center', valign: 'bottom' });
}

// 24 - Funnel Infographic
function slide24(p) {
  const s = newSlide(p, 24);
  head(s, 'Funnel Infoghraphic', { x: 3.219, y: 0.774, w: 6.896, h: 0.767, align: 'center', valign: 'bottom' });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', { x: 3.569, y: 1.567, w: 6.204, h: 0.341, align: 'center' });
  const tiers = [
    { y: 2.443, x: 2.353, w: 3.986, h: 1.044, rim: C.accentDk, body: C.accent, bar: 5.906, barW: 4.19, barFill: C.accentDk, pctX: 10.197, pct: '80' },
    { y: 3.557, x: 2.875, w: 2.942, h: 1.019, rim: 'A2A2A2', body: C.gray1, bar: 5.371, barW: 4.19, barFill: '6C6C6C', pctX: 9.662, pct: '75' },
    { y: 4.615, x: 3.395, w: 1.900, h: 1.014, rim: '8F8F8F', body: C.gray2, bar: 4.863, barW: 4.193, barFill: '8F8F8F', pctX: 9.157, pct: '70' },
    { y: 5.741, x: 3.917, w: 0.858, h: 0.884, rim: '7C7C7C', body: C.gray3, bar: 4.345, barW: 4.193, barFill: '7C7C7C', pctX: 8.639, pct: '80' }
  ];
  tiers.forEach((t, i) => {
    // caption bar behind the funnel tier
    rect(s, t.bar, t.y + 0.16, t.barW, 0.884, { color: t.barFill });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod',
      { x: [6.466, 6.007, 5.51, 4.89][i], y: [2.728, 3.799, 4.869, 5.981][i], w: 3.554, h: 0.604, color: C.white });
    s.addText([{ text: t.pct }, { text: '%', options: { superscript: true } }],
      { x: t.pctX, y: t.y + 0.31, w: 0.813, h: 0.584, valign: 'middle', fontFace: 'Lexend Deca', fontSize: 24, color: C.gray5, lineSpacingMultiple: 1.3 });
    // funnel tier = elliptical rim + tapering trapezoid
    const nw = i < 3 ? tiers[i + 1].w : t.w * 0.55;
    const bh = t.h - 0.16;
    s.addShape('ellipse', { x: t.x, y: t.y, w: t.w, h: 0.32, fill: { color: t.rim }, line: NOLINE });
    s.addShape('custGeom', {
      x: t.x, y: t.y + 0.16, w: t.w, h: bh, fill: { color: t.body }, line: NOLINE,
      points: [{ x: 0, y: 0 }, { x: t.w, y: 0 }, { x: t.w / 2 + nw / 2, y: bh }, { x: t.w / 2 - nw / 2, y: bh }, { close: true }]
    });
    icon(s, t.x + t.w / 2 - 0.21, t.y + 0.36, 0.42, C.white);
  });
}

// 25 - Gear Infographic Slide
function slide25(p) {
  const s = newSlide(p, 25);
  head(s, 'Gear Infographic Slide', { x: 2.957, y: 0.717, w: 7.053, h: 0.774, fontSize: 40, align: 'center', lineSpacingMultiple: 1 });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, sed do'),
    { x: 3.399, y: 1.61, w: 6.57, h: 0.604, align: 'center' });
  // dashed orbit arcs behind the gears
  s.addShape('arc', { x: 7.128, y: 3.657, w: 1.448, h: 0.909, line: { color: C.gray2, width: 1, dashType: 'dash' }, angleRange: [180, 0], rotate: 75 });
  s.addShape('arc', { x: 5.555, y: 5.825, w: 1.235, h: 1.235, line: { color: C.gray2, width: 1, dashType: 'dash' }, angleRange: [300, 90] });
  s.addShape('arc', { x: 3.961, y: 4.082, w: 1.176, h: 0.891, line: { color: C.gray2, width: 1, dashType: 'dash' }, angleRange: [20, 200] });
  const gears = [
    { x: 5.959, y: 2.644, sz: 1.807, tooth: C.accent, hub: C.accentDk, rot: 15 },
    { x: 4.487, y: 2.932, sz: 1.458, tooth: C.gray3, hub: '7C7C7C', rot: 0 },
    { x: 4.424, y: 4.77, sz: 1.807, tooth: C.gray2, hub: '8F8F8F', rot: 20 },
    { x: 6.236, y: 4.488, sz: 2.227, tooth: C.gray1, hub: 'C2C2C2', rot: 10 }
  ];
  gears.forEach(g => {
    s.addShape('gear9', { x: g.x, y: g.y, w: g.sz, h: g.sz, fill: { color: g.tooth }, line: NOLINE, rotate: g.rot });
    const hub = g.sz * 0.555;
    s.addShape('ellipse', { x: g.x + (g.sz - hub) / 2, y: g.y + (g.sz - hub) / 2, w: hub, h: hub, fill: { color: g.hub }, line: NOLINE });
    icon(s, g.x + g.sz / 2 - hub * 0.22, g.y + g.sz / 2 - hub * 0.22, hub * 0.44, C.white);
  });
  const stats = [
    { num: '40K', color: C.dark, nx: 2.048, tx: 1.327, y: 2.728, align: 'right', bx: 0.55, by: 3.694, bw: 3.325 },
    { num: '60+', color: C.accent, nx: 8.373, tx: 8.373, y: 2.591, align: 'left', bx: 8.392, by: 3.509, bw: 3.107 },
    { num: '70+', color: C.dark, nx: 2.123, tx: 1.402, y: 5.252, align: 'right', bx: 0.603, by: 6.203, bw: 3.325 },
    { num: '80K', color: C.dark, nx: 8.824, tx: 8.824, y: 5.252, align: 'left', bx: 8.872, by: 6.203, bw: 3.107 }
  ];
  stats.forEach(t => {
    s.addText(t.num, { x: t.nx, y: t.y, w: 1.827, h: 0.635, align: t.align, fontFace: BODY, fontSize: 28, color: t.color, valign: 'top', lineSpacingMultiple: 1.2 });
    s.addText('Your Text Here', { x: t.tx, y: t.y + 0.566, w: 2.548, h: 0.401, align: t.align, fontFace: BODY, fontSize: 16, bold: true, color: t.color, valign: 'top', lineSpacingMultiple: 1.2 });
    body(s, LOREM.tiny, { x: t.bx, y: t.by, w: t.bw, h: 0.604, align: t.align });
  });
}

// 26 - Infographic Slide (four chevron arrows)
function slide26(p) {
  const s = newSlide(p, 26);
  rect(s, 1.16, 3.398, 9.932, 0.269, { color: C.accent });
  const steps = [
    { x: 1.16, arrow: C.accent, tip: C.accentDk },
    { x: 4.065, arrow: C.gray1, tip: '363636' },
    { x: 6.97, arrow: C.gray2, tip: '8F8F8F' },
    { x: 9.875, arrow: C.gray3, tip: '7C7C7C' }
  ];
  steps.forEach((st, i) => {
    // banner slanting down-right, ending in an arrow point
    s.addShape('custGeom', {
      x: st.x, y: 3.017, w: 2.1, h: 1.67, fill: { color: st.arrow }, line: NOLINE,
      points: [{ x: 0.30, y: 0 }, { x: 1.42, y: 0 }, { x: 2.10, y: 1.42 }, { x: 2.05, y: 1.67 },
        { x: 1.53, y: 1.67 }, { x: 1.24, y: 1.42 }, { close: true }]
    });
    s.addShape('triangle', { x: st.x, y: 3.017, w: 0.547, h: 0.382, fill: { color: st.tip }, line: NOLINE });
    icon(s, st.x + 1.44, 4.0, 0.42, C.white);
    s.addText((i + 1) + '. Tittle Text Here', { x: [1.077, 4.148, 7.075, 10.048][i], y: i === 0 ? 5.061 : 5.04, w: 2.348, h: 0.37, fontFace: BODY, fontSize: 16, bold: true, color: C.gray5 });
    body(s, LOREM.two, { x: [1.077, 4.148, 7.075, 10.048][i], y: i === 0 ? 5.427 : 5.407, w: 2.693, h: 1.129 });
  });
  head(s, 'Infographic Slide', { x: 2.695, y: 0.851, w: 7.943, h: 0.774, fontSize: 40, align: 'center', lineSpacingMultiple: 1 });
  body(s, lead('dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor adipiscing elit, sed do'),
    { x: 3.399, y: 1.658, w: 6.57, h: 0.604, align: 'center' });
}

/** Draw a map region from normalised outline points inside a bounding box. */
function mapShape(s, box, pts, color) {
  s.addShape('custGeom', {
    x: box.x, y: box.y, w: box.w, h: box.h, fill: { color }, line: { color: C.white, width: 0.75 },
    points: pts.map(pt => ({ x: +(box.w * pt[0]).toFixed(3), y: +(box.h * pt[1]).toFixed(3) })).concat([{ close: true }])
  });
}
/** Regions of each map, painted back to front: [fill colour, outline points]. */
const BELGIUM = [
  [C.gray2, [[0.213, 0.063], [0.220, 0.148], [0.231, 0.230], [0.241, 0.311], [0.254, 0.354], [0.084, 0.396],
    [0.047, 0.354], [0.019, 0.311], [0.024, 0.272], [0.006, 0.230], [0.028, 0.187], [0.075, 0.148], [0.131, 0.106]]],
  [C.gray1, [[0.588, 0.003], [0.659, 0.034], [0.690, 0.095], [0.707, 0.158], [0.675, 0.190], [0.642, 0.219],
    [0.537, 0.251], [0.435, 0.219], [0.442, 0.190], [0.470, 0.158], [0.461, 0.127], [0.494, 0.063], [0.485, 0.034]]],
  [C.gray2, [[0.442, 0.069], [0.461, 0.148], [0.438, 0.187], [0.420, 0.269], [0.403, 0.309], [0.399, 0.348],
    [0.287, 0.388], [0.261, 0.348], [0.237, 0.269], [0.239, 0.227], [0.224, 0.148], [0.267, 0.108]]],
  [C.gray1, [[0.672, 0.227], [0.657, 0.272], [0.679, 0.317], [0.668, 0.359], [0.668, 0.383], [0.381, 0.404],
    [0.373, 0.383], [0.403, 0.338], [0.403, 0.317], [0.425, 0.272], [0.444, 0.248]]],
  [C.gray2, [[0.815, 0.351], [0.903, 0.393], [0.940, 0.435], [0.983, 0.517], [0.987, 0.559], [0.972, 0.602],
    [0.938, 0.644], [0.933, 0.686], [0.907, 0.644], [0.838, 0.602], [0.718, 0.559], [0.700, 0.517], [0.657, 0.478],
    [0.640, 0.435], [0.644, 0.393]]],
  ['E4E4E4', [[0.582, 0.348], [0.642, 0.401], [0.552, 0.456], [0.537, 0.512], [0.534, 0.565], [0.463, 0.617],
    [0.487, 0.728], [0.496, 0.781], [0.429, 0.728], [0.418, 0.673], [0.416, 0.617], [0.295, 0.565], [0.282, 0.512],
    [0.192, 0.456], [0.183, 0.401]]],
  [C.accent, [[0.772, 0.108], [0.789, 0.145], [0.858, 0.182], [0.849, 0.219], [0.830, 0.290], [0.802, 0.327],
    [0.793, 0.364], [0.685, 0.401], [0.672, 0.364], [0.679, 0.327], [0.664, 0.290], [0.672, 0.253], [0.647, 0.219],
    [0.696, 0.182], [0.703, 0.145]]],
  [C.accent, [[0.634, 0.430], [0.675, 0.501], [0.815, 0.573], [0.903, 0.644], [0.866, 0.712], [0.836, 0.784],
    [0.834, 0.855], [0.871, 0.926], [0.767, 0.997], [0.722, 0.926], [0.606, 0.855], [0.603, 0.784], [0.485, 0.712],
    [0.487, 0.644], [0.534, 0.573], [0.530, 0.501]]]
];
const UK = [
  ['F2F2F2', [[0.535, 0.002], [0.596, 0.100], [0.669, 0.195], [0.759, 0.293], [0.814, 0.390], [0.788, 0.486],
    [0.701, 0.584], [0.849, 0.679], [0.590, 0.777], [0.593, 0.679], [0.407, 0.584], [0.413, 0.486], [0.538, 0.390],
    [0.515, 0.293], [0.506, 0.195], [0.456, 0.100]]],
  [C.gray1, [[0.884, 0.445], [0.980, 0.473], [0.994, 0.501], [0.994, 0.529], [0.980, 0.557], [0.936, 0.614],
    [0.895, 0.642], [0.831, 0.670], [0.724, 0.642], [0.718, 0.614], [0.709, 0.586], [0.706, 0.557], [0.738, 0.529],
    [0.727, 0.501], [0.817, 0.473], [0.849, 0.445]]],
  [C.gray2, [[0.567, 0.577], [0.576, 0.653], [0.593, 0.692], [0.581, 0.729], [0.558, 0.768], [0.349, 0.807],
    [0.331, 0.844], [0.145, 0.883], [0.137, 0.844], [0.169, 0.807], [0.227, 0.768], [0.273, 0.729], [0.416, 0.692],
    [0.459, 0.653], [0.459, 0.616]]],
  [C.accent, [[0.453, 0.100], [0.468, 0.143], [0.500, 0.187], [0.494, 0.230], [0.477, 0.273], [0.523, 0.315],
    [0.547, 0.358], [0.535, 0.401], [0.483, 0.445], [0.407, 0.401], [0.404, 0.358], [0.410, 0.315], [0.436, 0.273],
    [0.366, 0.230], [0.346, 0.187], [0.381, 0.143]]]
];

// 27 - Belgium Map Infographic
function slide27(p) {
  const s = newSlide(p, 27);
  const box = { x: 6.474, y: 1.349, w: 6.442, h: 5.261 };
  BELGIUM.forEach(([color, pts]) => mapShape(s, box, pts, color));
  s.addText('Belgium Map\nInfographic', { x: 1.119, y: 1.37, w: 5.721, h: 1.447, fontFace: HEAD, fontSize: 40, color: C.gray5, valign: 'middle', lineSpacingMultiple: 1.1 });
  s.addText('Your Text Here', { x: 1.119, y: 3.329, w: 3.385, h: 0.404, fontFace: BODY, fontSize: 18, bold: true, color: C.gray5 });
  body(s, LOREM.two, { x: 1.124, y: 3.72, w: 4.715, h: 0.604 });
  [{ x: 1.126, bx: 1.117, num: '350K', color: C.accent }, { x: 4.97, bx: 4.961, num: '60+', color: C.gray3 }].forEach(t => {
    s.addText(t.num, { x: t.x, y: 5.252, w: 1.127, h: 0.438, fontFace: BODY, fontSize: 20, bold: true, color: t.color });
    s.addText('Your Text Here', { x: t.bx, y: 5.656, w: 2.548, h: 0.401, fontFace: BODY, fontSize: 16, bold: true, color: t.color, lineSpacingMultiple: 1.2 });
    body(s, LOREM.tiny, { x: t.bx, y: t.x < 2 ? 6.083 : 6.11, w: 2.978, h: 0.604 });
  });
}

// 28 - United Kingdom Maps
function slide28(p) {
  const s = newSlide(p, 28);
  const box = { x: 6.391, y: 1.004, w: 4.782, h: 6.397 };
  UK.forEach(([color, pts]) => mapShape(s, box, pts, color));
  s.addText('United Kingdom Maps', { x: 1.851, y: 1.038, w: 6.569, h: 0.707, fontFace: HEAD, fontSize: 40, color: C.gray5, valign: 'bottom', lineSpacingMultiple: 0.9 });
  [{ y: 2.129, pct: 75, color: C.accent, tc: C.accent }, { y: 3.741, pct: 82, color: C.gray1, tc: C.gray5 },
    { y: 5.291, pct: 85, color: C.gray2, tc: C.gray5 }].forEach((g, i) => {
    ring(s, 1.851, g.y + 0.056, 1.098, g.pct, g.color, String(g.pct), { rest: 'F2F2F2', hole: 0.62, fontSize: 20 });
    body(s, 'Feature Details', { x: 3.186, y: g.y, w: 2.792, h: 0.381, fontSize: 14, bold: true, color: g.tc });
    body(s, LOREM.tiny, { x: [3.186, 3.18, 3.192][i], y: g.y + 0.41, w: 3.3, h: 0.604 });
  });
}

// 29 - Contact Us
function slide29(p) {
  const s = newSlide(p, 29);
  rect(s, 0, 0, 6.161, 7.5, { color: C.accent });
  s.addShape('teardrop', { x: 6.575, y: 2.026, w: 0.324, h: 0.324, fill: { color: C.accent }, line: NOLINE, rotate: 135 });
  s.addShape('ellipse', { x: 6.659, y: 2.122, w: 0.156, h: 0.156, fill: { color: C.white }, line: NOLINE });
  s.addShape('roundRect', { x: 5.903, y: 2.562, w: 1.669, h: 0.372, fill: { color: C.accent }, line: NOLINE, rectRadius: 0.05 });
  s.addText('Your Text Here', { x: 5.903, y: 2.562, w: 1.669, h: 0.372, align: 'center', valign: 'middle', fontFace: 'Lexend Deca', fontSize: 14, color: C.white });
  head(s, 'Contact Us', { x: 8.437, y: 1.564, w: 3.966, h: 0.841, lineSpacingMultiple: 1 });
  s.addText('Your Text Here', { x: 8.481, y: 2.838, w: 3.385, h: 0.404, fontFace: BODY, fontSize: 18, color: C.gray5 });
  body(s, LOREM.tiny, { x: 8.481, y: 3.33, w: 4.186, h: 0.604 });
  const contacts = ['666 5634 7776', 'Company Street 831, New York', 'mail@company.com', 'www.companysite.com'];
  contacts.forEach((t, i) => {
    const y = 4.411 + i * 0.373, cx = 8.62, cy = y + 0.09;
    if (i === 0) { // handset
      s.addShape('roundRect', { x: cx, y: cy, w: 0.075, h: 0.15, fill: { color: C.accent }, line: NOLINE, rectRadius: 0.03, rotate: 330 });
    } else if (i === 1) { // house
      s.addShape('triangle', { x: cx - 0.015, y: cy, w: 0.17, h: 0.08, fill: { color: C.accent }, line: NOLINE });
      s.addShape('rect', { x: cx + 0.015, y: cy + 0.075, w: 0.11, h: 0.08, fill: { color: C.accent }, line: NOLINE });
    } else if (i === 2) { // envelope
      s.addShape('rect', { x: cx - 0.015, y: cy + 0.02, w: 0.17, h: 0.115, fill: { color: C.accent }, line: NOLINE });
      s.addShape('triangle', { x: cx - 0.015, y: cy + 0.02, w: 0.17, h: 0.075, fill: { color: C.white }, line: NOLINE, rotate: 180 });
    } else { // globe
      s.addShape('ellipse', { x: cx - 0.01, y: cy + 0.005, w: 0.155, h: 0.155, fill: { color: C.accent }, line: NOLINE });
      s.addShape('ellipse', { x: cx + 0.038, y: cy + 0.005, w: 0.06, h: 0.155, fill: NOLINE, line: { color: C.white, width: 0.75 } });
    }
    body(s, t, { x: 8.846, y, w: 3.084, h: 0.326, lineSpacingMultiple: 1.2 });
  });
}

// 30 - Thank You
function slide30(p) {
  const s = newSlide(p, 30);
  gradientV(s, 0, 0, 13.333, 7.5, [[0, 'F9F9F9'], [0.5, 'EBDFDA'], [1, 'D3B4A3']], 60);
  s.addText('Thank You', { x: 1.263, y: 2.861, w: 10.973, h: 1.836, align: 'center', valign: 'middle', fontFace: HEAD, fontSize: 96, color: C.white });
  s.addText('For Watching Our Presentation', {
    x: 3.381, y: 4.355, w: 6.737, h: 0.406, align: 'center', fontFace: BODY, fontSize: 16, color: C.white,
    charSpacing: 3, lineSpacingMultiple: 1.2
  });
}

/* -------------------------------------------------------------------- main */
function build() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'GAMISH', width: 13.333, height: 7.5 });
  pres.layout = 'GAMISH';
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pres.title = 'Gamish - Home Interior Presentation Template';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
    slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30]
    .forEach(fn => fn(pres));

  return pres.writeFile({ fileName: path.join(__dirname, OUT_NAME) });
}

const OUT_NAME = '09d49f47-ff8c-4200-873c-651f1ea2dede_grok_final.pptx';
build().then(f => console.log('wrote', f));
