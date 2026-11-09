/**
 * GreenEats pitch deck — rebuilt with pptxgenjs.
 * Slide canvas: 26.667in x 15in.
 * Raster photography in the source deck is replaced by flat colour placeholders.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const C = {
  lime: 'D6E35C',
  green: '639B2C',
  ink: '141414',
  black: '000000',
  white: 'FFFFFF',
  grey: '929292',
  slate: '504D51',
  body: '5E5E5E',
  mist: 'DFE4E4',
  pale: 'E6E9EA',
  silver: 'D6D6D6',
  photo: 'C4D4DC',   // stand-ins for the deck's photographs
  photoB: 'D2E3EB',
  photoC: 'D3E4ED'
};

const F = { reg: 'Inter', light: 'Inter Light', semi: 'Inter SemiBold' };

/* Recurring run styles */
const S = {
  eyebrow: { fontFace: F.reg, fontSize: 16, color: C.grey },
  micro: { fontFace: F.reg, fontSize: 16, color: C.grey },
  small: { fontFace: F.light, fontSize: 22, color: C.body, lineSpacingMultiple: 1.2 },
  lead: { fontFace: F.light, fontSize: 32, color: C.slate },
  head: { fontFace: F.semi, fontSize: 48, color: C.black },
  display: { fontFace: F.reg, fontSize: 120, bold: true, color: C.black },
  hero: { fontFace: F.reg, fontSize: 300, bold: true, color: C.ink }
};

/** Width/position of the middle breadcrumb, which varies with the topic name. */
const CRUMB = { med: [3.936, 2.465], wide: [3.551, 3.234], team: [3.931, 2.475],
  fin: [3.871, 2.763], narrow: [4.403, 1.809] };

/* ------------------------------------------------------------------ helpers */
const TEXT_INSET = 4;     // pt — matches the 50800 EMU inset used throughout
const SHAPE_INSET = 0;

/** Text box. `st` is one of the S.* presets (or any pptxgenjs text option bag). */
function txt(s, text, x, y, w, h, st, extra) {
  s.addText(text, Object.assign({ x, y, w, h, margin: TEXT_INSET, valign: 'middle' }, st, extra));
}

/** Filled rectangle. */
function rect(s, x, y, w, h, color, transparency) {
  s.addShape('rect', { x, y, w, h, fill: { color, transparency: transparency || 0 } });
}

/** Outlined (unfilled) rectangle. */
function frame(s, x, y, w, h, color) {
  s.addShape('rect', { x, y, w, h, fill: { type: 'none' }, line: { color, width: 2 } });
}

/** Straight rule; pass w = 0 for a vertical one. */
function rule(s, x, y, w, h, color) {
  s.addShape('line', { x, y, w, h, line: { color, width: 2 } });
}

/** Numbered circle badge. */
function badge(s, n, x, y, d, fill, color, size) {
  txt(s, n, x, y, d, d, { shape: 'ellipse', fill: { color: fill }, fontFace: F.reg,
    fontSize: size, bold: true, color: color, align: 'center', margin: SHAPE_INSET });
}

/** Flat block standing in for a photograph. `shape` is 'rect' or 'ellipse'. */
function photo(s, x, y, w, h, opt) {
  const o = opt || {};
  s.addShape(o.shape || 'rect', {
    x, y, w, h,
    fill: { color: o.color || C.photo, transparency: o.transparency || 0 },
    altText: '[image]'
  });
}

/** The "Chapter NN / Topic / Green Eats®" strip every content slide carries. */
function crumbs(s, chapter, topic, slot, color) {
  const st = Object.assign({}, S.eyebrow, { color: color || C.grey });
  txt(s, chapter, 1.284, 1.387, 1.729, 0.389, st);
  txt(s, topic, slot[0], 1.387, slot[1], 0.389, st);
  txt(s, 'Green Eats\u00AE', 7.323, 1.387, 1.729, 0.389, st);
}

/**
 * Path helper for custGeom: `spec` rows are either [x, y] line/move points or
 * [x1, y1, x2, y2, x, y] cubic segments, all as fractions of the shape box.
 */
function pathPoints(spec, w, h) {
  const pts = spec.map(function (p, i) {
    if (p.length === 6) {
      return { x: p[4] * w, y: p[5] * h,
        curve: { type: 'cubic', x1: p[0] * w, y1: p[1] * h, x2: p[2] * w, y2: p[3] * h } };
    }
    return { x: p[0] * w, y: p[1] * h, moveTo: i === 0 };
  });
  pts.push({ close: true });
  return pts;
}

/** Rounded speech balloon with a tail on the lower left (title + vision slides). */
function balloon(s, x, y, w, h, color) {
  const spec = [
    [0.104, 0.000], [0.070, 0.000, 0.042, 0.055, 0.042, 0.122], [0.042, 0.806],
    [0.000, 1.000], [0.108, 0.949], [0.938, 0.949],
    [0.972, 0.949, 1.000, 0.894, 1.000, 0.827], [1.000, 0.122],
    [1.000, 0.055, 0.972, 0.000, 0.938, 0.000], [0.104, 0.000]
  ];
  s.addShape('custGeom', { x, y, w, h, fill: { color }, points: pathPoints(spec, w, h) });
}

/** GreenEats logo: green disc with a lime tree cut into it. */
function logoMark(s, x, y, d) {
  s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: C.green } });
  s.addShape('ellipse', { x: x + 0.27 * d, y: y + 0.20 * d, w: 0.46 * d, h: 0.40 * d, fill: { color: C.lime } });
  rect(s, x + 0.44 * d, y + 0.44 * d, 0.12 * d, 0.34 * d, C.lime);
}

/** Solid arrow outline, drawn pointing up inside its box. */
const ARROW = [[0.501, 0], [0, 0.379], [0.150, 0.494], [0.394, 0.308], [0.394, 1],
  [0.608, 1], [0.608, 0.308], [0.850, 0.494], [1, 0.379], [0.501, 0]];

/** Slim arrow glyph on the team cards; `deg` 0 = up, 90 = right. */
function arrowGlyph(s, x, y, w, h, deg, color) {
  s.addShape('custGeom', { x, y, w, h, rotate: deg, fill: { color }, points: pathPoints(ARROW, w, h) });
}

/** Turn an array of strings into bulleted paragraphs. */
function bullets(items) {
  return items.map(function (t) {
    return { text: t, options: { breakLine: true, bullet: { characterCode: '2022', indent: 22 } } };
  });
}

/**
 * Inner marks for the icon-library sheets. Each draws inside the unit box
 * (cx, cy = centre; u = the mark's full width/height in inches).
 */
const STROKE = 0.1;                      // common line weight of the icon set
const MARKS = {
  bars: function (s, cx, cy, u) {
    [-0.32, 0, 0.32].forEach(function (k) {
      rect(s, cx - u / 2, cy + k * u - STROKE / 2, u, STROKE, C.black);
    });
  },
  list: function (s, cx, cy, u) {
    [-0.32, 0, 0.32].forEach(function (k) {
      s.addShape('ellipse', { x: cx - u / 2, y: cy + k * u - STROKE / 2, w: STROKE, h: STROKE, fill: { color: C.black } });
      rect(s, cx - u / 2 + 0.22 * u, cy + k * u - STROKE / 2, 0.78 * u, STROKE, C.black);
    });
  },
  box: function (s, cx, cy, u) {
    frame(s, cx - u / 2, cy - u / 2, u, u, C.black);
  },
  boxFilled: function (s, cx, cy, u) {
    rect(s, cx - u / 2, cy - u / 2, u, u, C.black);
  },
  dot: function (s, cx, cy, u) {
    s.addShape('ellipse', { x: cx - u / 4, y: cy - u / 4, w: u / 2, h: u / 2, fill: { color: C.black } });
  },
  ring: function (s, cx, cy, u) {
    s.addShape('ellipse', { x: cx - u / 2, y: cy - u / 2, w: u, h: u,
      fill: { type: 'none' }, line: { color: C.black, width: STROKE * 72 } });
  },
  plus: function (s, cx, cy, u) {
    rect(s, cx - u / 2, cy - STROKE / 2, u, STROKE, C.black);
    rect(s, cx - STROKE / 2, cy - u / 2, STROKE, u, C.black);
  },
  cross: function (s, cx, cy, u) {
    [45, -45].forEach(function (a) {
      s.addShape('rect', { x: cx - u / 2, y: cy - STROKE / 2, w: u, h: STROKE, rotate: a, fill: { color: C.black } });
    });
  },
  minus: function (s, cx, cy, u) {
    rect(s, cx - u / 2, cy - STROKE / 2, u, STROKE, C.black);
  },
  tri: function (s, cx, cy, u) {
    s.addShape('triangle', { x: cx - u / 2, y: cy - u / 2, w: u, h: u, fill: { color: C.black } });
  }
};

['up', 'right', 'down', 'left', 'upRight', 'downRight', 'downLeft', 'upLeft']
  .forEach(function (dir, i) {
    const deg = i < 4 ? i * 90 : 45 + (i - 4) * 90;
    MARKS[dir] = function (s, cx, cy, u) { arrowGlyph(s, cx - u / 2, cy - u / 2, u, u, deg, C.black); };
  });

/** One outline-circle pictogram: heavy ring plus an inner mark. */
function ringIcon(s, x, y, d, mark) {
  s.addShape('ellipse', {
    x: x + STROKE / 2, y: y + STROKE / 2, w: d - STROKE, h: d - STROKE,
    fill: { type: 'none' }, line: { color: C.black, width: STROKE * 72 }
  });
  MARKS[mark](s, x + d / 2, y + d / 2, d * 0.45);
}

/** Grid of ring pictograms; each row supplies its own column positions. */
function iconGrid(s, d, rows) {
  rows.forEach(function (r) {
    r.marks.forEach(function (m, c) { ringIcon(s, r.xs[c], r.y, d, m); });
  });
}

/* ------------------------------------------------------------------- slides */

// 1 — Title
function slide01(pres) {
  const s = pres.addSlide();
  s.background = { color: C.lime };
  photo(s, 17.615, 0, 9.052, 15.0);
  txt(s, 'GreenEats', 1.284, 3.016, 13.511, 8.682, S.hero, { lineSpacingMultiple: 0.7 });
  balloon(s, 9.661, 7.072, 4.76, 2.405, C.green);
  txt(s, 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum.',
    10.215, 7.589, 3.921, 1.371, S.small, { color: C.white });
  rect(s, 0, 13.712, 17.62, 1.288, C.ink);
  txt(s, '\u00A9 2077 Provue. All rights reserved.', 1.284, 14.162, 3.801, 0.389, S.micro);
  logoMark(s, 1.286, 1.108, 0.945);
  txt(s, [
    { text: 'Present by ', options: { fontFace: F.light } },
    { text: 'Mark Anderson', options: { fontFace: F.reg, bold: true } }
  ], 1.485, 11.331, 6.447, 0.653, { fontSize: 32, color: C.slate });
}

// 2 — Introduction
function slide02(pres) {
  const s = pres.addSlide();
  photo(s, 1.286, 5.529, 13.889, 8.363);
  crumbs(s, 'Chapter 01', 'Executive Summary', CRUMB.med);
  txt(s, 'Introduction', 1.284, 2.922, 13.511, 2.125, S.display);
  txt(s, [
    { text: 'GreenEats is a sustainable food delivery service dedicated to providing fresh, locally sourced, and eco-friendly meals to our customers. ', options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: 'Our mission is to make sustainable eating convenient, delicious, and accessible to all.' }
  ], 16.528, 5.547, 7.907, 4.444, S.lead);
}

// 3 — Company Overview divider
function slide03(pres) {
  const s = pres.addSlide();
  s.background = { color: C.green };
  photo(s, 0, 0, 23.053, 15.0, { transparency: 85 });
  txt(s, 'Green Eats\u00AE', 1.284, 1.387, 1.729, 0.389, S.eyebrow, { color: C.white });
  txt(s, 'Company Overview', 1.286, 4.705, 23.803, 9.186, S.hero,
    { color: C.white, lineSpacingMultiple: 0.8 });
}

// 4 — Our Vision
function slide04(pres) {
  const s = pres.addSlide();
  txt(s, 'To create a world where                  sustainable food           choices are accessible, convenient,               and enjoyable for everyone, leading to a healthier planet        and happier communities.',
    1.284, 3.078, 24.099, 10.212, { fontFace: F.reg, fontSize: 100, bold: true, color: C.black });
  photo(s, 13.32, 4.805, 2.465, 1.484);
  photo(s, 17.738, 3.265, 2.465, 1.484);
  photo(s, 18.086, 6.609, 3.713, 1.484);
  crumbs(s, 'Chapter 02', 'Our Vision', CRUMB.med);
  txt(s, 'GreenEats is a sustainable food delivery service dedicated to providing fresh, locally sourced, and eco-friendly meals to our customers.',
    10.969, 12.016, 10.421, 0.927, S.small);
  balloon(s, 13.616, 10.092, 1.656, 1.3, C.green);
  txt(s, '\uD83C\uDF0D', 13.626, 10.219, 1.729, 0.92,
    { fontFace: F.semi, fontSize: 48, color: C.slate, align: 'center' });
}

// 5 — Our Values
function slide05(pres) {
  const s = pres.addSlide();
  crumbs(s, 'Chapter 02', 'Our Values', CRUMB.med);
  txt(s, 'Nourishing People, Preserve the Planet', 1.286, 3.078, 16.424, 4.152, S.display);
  txt(s, "GreenEats embraces sustainability as a core principle, offering numerous advantages for our customers, the environment, and the community. Here's why choosing GreenEats is a game-changer",
    17.62, 5.147, 7.765, 1.772, S.small);

  const cards = [
    { x: 1.286, fill: C.green, titleX: 1.703, title: 'Sustainability', ink: C.white,
      bodyX: 1.751, bodyY: 11.131, bodyW: 6.416,
      body: 'We prioritize sustainable practices in every aspect of our business, from ingredient sourcing to packaging and delivery' },
    { x: 8.628, fill: C.lime, titleX: 9.049, title: 'Quality', ink: C.black, bodyInk: C.body,
      bodyX: 9.049, bodyY: 11.131, bodyW: 6.503,
      body: 'We are committed to delivering high-quality, flavorful meals that nourish our customers and support their well-being.' },
    { x: 15.974, fill: C.mist, titleX: 16.391, title: 'Convenience', ink: C.slate, bodyInk: C.body,
      bodyX: 16.391, bodyY: 11.085, bodyW: 6.503,
      body: 'We aim to make sustainable eating easy and accessible, providing hassle-free delivery and diverse menu options.' }
  ];
  cards.forEach(function (c) {
    rect(s, c.x, 8.518, 7.347, 4.368, c.fill);
    txt(s, c.title, c.titleX, 8.935, 4.703, 0.931, S.head, { color: c.ink });
    txt(s, c.body, c.bodyX, c.bodyY, c.bodyW, 1.339, S.small, { color: c.bodyInk || C.white });
  });
}

// 6 — Problem Statement
function slide06(pres) {
  const s = pres.addSlide();
  crumbs(s, 'Chapter 03', 'Problem Statement', CRUMB.med);
  txt(s, 'Problem Statement', 1.286, 2.922, 18.177, 2.132, S.display);

  const cards = [
    { x: 1.286, numX: 1.842, n: '01', titleX: 1.842, titleY: 9.313, titleW: 6.654, titleH: 1.75,
      title: 'Excessive Packaging Waste', bodyX: 1.751,
      body: 'Single-use plastic containers, cutlery, and excessive packaging contribute to landfill waste and pollution.' },
    { x: 9.453, numX: 10.009, n: '02', titleX: 10.099, titleY: 9.313, titleW: 6.872, titleH: 1.75,
      title: 'Processed \nMeals Domination', bodyX: 10.009,
      body: 'Fast food and processed meals dominate the market, leading to negative health impacts and unsustainable consumption patterns.' },
    { x: 17.615, numX: 18.176, n: '03', titleX: 18.254, titleY: 9.467, titleW: 6.654, titleH: 0.931,
      title: 'Carbon Emissions', bodyX: 18.176,
      body: 'Conventional food delivery services contribute to excessive packaging waste and carbon emissions from transportation.' }
  ];
  cards.forEach(function (c) {
    rect(s, c.x, 8.341, 7.765, 5.551, C.mist);
    badge(s, c.n, c.numX, 7.646, 1.389, C.lime, C.black, 32);
    txt(s, c.title, c.titleX, c.titleY, c.titleW, c.titleH, S.head);
    txt(s, c.body, c.bodyX, 11.922, 6.416, 1.339, S.small);
  });
}

// 7 — Problem statistics on green
function slide07(pres) {
  const s = pres.addSlide();
  s.background = { color: C.green };
  photo(s, 20.542, 10.634, 8.732, 8.732, { shape: 'ellipse' });
  photo(s, 14.327, 2.123, 4.719, 4.719, { shape: 'ellipse' });
  photo(s, 10.193, 12.09, 2.36, 2.36, { shape: 'ellipse' });
  crumbs(s, 'Chapter 03', 'Problem Statement', CRUMB.med, C.silver);
  txt(s, 'The current food delivery landscape presents several challenges, including environmental impact and the lack of healthy, sustainable options.',
    1.286, 3.31, 11.267, 1.736, { fontFace: F.semi, fontSize: 32, color: C.white });

  const stats = [
    { x: 8.563, y: 6.106, big: '9%', bigW: 4.577, capW: 6.431, capH: 1.736,
      cap: 'PLACEHOLDER' },
    { x: 14.947, y: 10.393, big: '14%', bigW: 4.577, capW: 5.285, capH: 1.194,
      cap: 'Carbon emissions from transportation account' },
    { x: 2.237, y: 9.967, big: '8 Mio', bigW: 4.577, capW: 6.121, capH: 1.194,
      cap: 'metric tons of plastic waste enter the oceans each year' },
    { x: 20.233, y: 6.842, big: '25%+', bigW: 5.148, capW: 4.577, capH: 1.194,
      cap: 'Sales of processed and packaged foods' }
  ];
  stats.forEach(function (t) {
    txt(s, t.big, t.x, t.y, t.bigW, 2.125, S.display, { color: C.white });
    txt(s, t.cap, t.x, t.y + 2.125, t.capW, t.capH, { fontFace: F.reg, fontSize: 32, color: C.white });
  });
}

// 8 — Market Analysis
function slide08(pres) {
  const s = pres.addSlide();
  crumbs(s, 'Chapter 04', 'Market Analysis', CRUMB.med);
  txt(s, 'Market Analysis', 1.286, 2.922, 15.531, 2.125, S.display);

  rect(s, 1.284, 6.193, 11.487, 1.389, C.pale);
  rect(s, 12.771, 6.193, 7.429, 1.389, C.silver);
  rect(s, 20.199, 6.193, 5.181, 1.389, C.lime);
  rule(s, 12.757, 6.193, 0, 3.216, C.pale);
  rule(s, 20.186, 6.193, 0, 3.216, C.silver);
  rule(s, 25.395, 6.193, 0, 3.216, C.black);

  txt(s, '60%', 8.597, 7.582, 4.146, 2.125, S.display, { color: C.pale });
  txt(s, '25%', 16.025, 7.582, 4.146, 2.125, S.display, { color: C.silver });
  txt(s, '15%', 21.234, 7.582, 4.146, 2.125, S.display, { color: C.green });
  txt(s, 'Our Market', 20.441, 6.561, 4.699, 0.653,
    { fontFace: F.semi, fontSize: 32, color: C.black, align: 'center' });
  txt(s, 'Understanding the target market and consumer trends is crucial to the success of GreenEats. Here is an overview of the market landscape and the demand for sustainable and healthy food options.',
    9.437, 10.784, 10.879, 2.278, S.lead);
  txt(s, 'Sources : Opta Research', 21.498, 12.589, 3.912, 0.472,
    { fontFace: F.reg, fontSize: 22, italic: true, color: C.body, align: 'right', lineSpacingMultiple: 1.2 });
}

// 9 — Solutions
function slide09(pres) {
  const s = pres.addSlide();
  photo(s, 9.799, 1.108, 13.889, 8.301);
  crumbs(s, 'Chapter 05', 'Our Solution', CRUMB.med);
  txt(s, 'Solutions', 1.284, 2.922, 10.108, 2.125, S.display);
  txt(s, 'We provide a solution for individuals who seek convenient, nutritious meals without compromising on their commitment to sustainability.',
    1.284, 6.815, 6.649, 2.819, S.lead);

  const steps = [
    { x: 9.799, n: '01', h: 2.569, label: 'Menu \nof Chef Crafted' },
    { x: 14.381, n: '02', h: 2.569, label: 'Sustainably Sourced Meals' },
    { x: 19.587, n: '03', h: 1.75, label: 'Organic Ingredients.' }
  ];
  steps.forEach(function (t) {
    badge(s, t.n, t.x, 9.849, 1.389, C.lime, C.black, 32);
    txt(s, t.label, t.x, 11.557, 4.101, t.h, S.head);
  });
}

// 10 — Our Business Model
function slide10(pres) {
  const s = pres.addSlide();
  photo(s, 0, 5.547, 9.051, 8.301);
  crumbs(s, 'Chapter 05', 'Our Solution', CRUMB.med);
  txt(s, 'Our Business Model', 1.284, 2.922, 19.28, 2.132, S.display);
  rect(s, 7.323, 7.279, 18.058, 5.665, C.mist);

  const lorem = 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat.';
  const cols = [
    { x: 8.127, titleY: 7.999, ruleY: 8.943, loremY: 9.235, listH: 2.206, title: 'Cost Structures',
      items: ['Ingredients Procurement', 'Meal Preparation', 'Packaging', 'Delivery Logistics', 'Operational Expenses'] },
    { x: 13.885, titleY: 7.992, ruleY: 8.936, loremY: 9.228, listH: 0.906, title: 'Key Partnerships',
      items: ['Local Farmers', 'Local Lawyer '] },
    { x: 19.5, titleY: 7.978, ruleY: 8.95, loremY: 9.214, listH: 1.339, title: 'Revenue Stream',
      items: ['Meal Sales', 'Subscription Plans', 'Customized Catering Services'] }
  ];
  cols.forEach(function (c) {
    txt(s, c.title, c.x, c.titleY, 5.076, 0.653, { fontFace: F.reg, fontSize: 32, bold: true, color: C.slate });
    rule(s, c.x, c.ruleY, 5.076, 0, C.black);
    txt(s, lorem, c.x, c.loremY, 5.076, 0.667, S.micro);
    txt(s, bullets(c.items), c.x, 10.04, 5.076, c.listH, S.small);
  });
}

// 11 — Marketing & Sales Strategy
function slide11(pres) {
  const s = pres.addSlide();
  const tiles = [
    [0, 3.997, C.photo], [5.872, 3.997, C.photoB], [11.743, 3.977, C.photo],
    [0, 9.588, C.photo], [5.872, 9.588, C.photoC], [11.743, 9.568, C.photo]
  ];
  tiles.forEach(function (t) { photo(s, t[0], t[1], 5.872, 5.432, { color: t[2] }); });
  crumbs(s, 'Chapter 05', 'Marketing & Sales Strategy', CRUMB.wide);

  const panels = [
    { x: 17.62, y: 3.997, capY: 4.656, numY: 6.646, num: '5690+',
      cap: 'Focuses on communicating our commitment to sustainability, health, and convenience' },
    { x: 17.615, y: 9.568, capY: 10.227, numY: 12.217, num: '40%',
      cap: 'We also leverage partnerships with local gyms, wellness centers, and sustainable living communities' }
  ];
  panels.forEach(function (p) {
    rect(s, p.x, p.y, 9.046, 5.432, C.mist);
    txt(s, p.cap, 18.305, p.capY, 7.676, 1.736, S.lead);
    txt(s, p.num, 18.305, p.numY, 6.385, 2.125, S.display);
  });
}

// 12 — Competitive advantages (four green tiles over photos)
function slide12(pres) {
  const s = pres.addSlide();
  const tiles = [
    { x: 1.284, tintX: 1.286, tintY: 5.547, labelX: 1.647, label: 'Diverse Menu Options' },
    { x: 7.385, tintX: 7.385, tintY: 5.547, labelX: 7.748, label: 'Sustainable Sourcing' },
    { x: 13.483, tintX: 13.488, tintY: 5.529, labelX: 13.847, label: 'Seasonality and Freshness' },
    { x: 19.584, tintX: 19.58, tintY: 5.529, labelX: 19.948, label: 'We Support Biodiversity' }
  ];
  tiles.forEach(function (t) {
    photo(s, t.x, 5.529, 5.803, 9.471);
    rect(s, t.tintX, t.tintY, 5.799, 9.453, C.green, 20);
    txt(s, t.label, t.labelX, 5.907, 5.076, 1.75, S.head, { color: C.white });
  });
  crumbs(s, 'Chapter 05', 'Competitive Advantages', CRUMB.wide);
  txt(s, "At GreenEats, we take pride in our menu offerings, which prioritize sustainability, nutrition, and flavor. Here's what sets our menu apart",
    1.286, 2.784, 10.351, 1.736, S.lead);
}

// 13 — Operations
function slide13(pres) {
  const s = pres.addSlide();
  photo(s, 1.284, 2.764, 10.769, 11.127);
  photo(s, 18.86, 10.093, 6.521, 3.799);
  photo(s, 12.192, 10.092, 6.521, 3.799);
  rect(s, 18.86, 10.092, 6.521, 3.8, C.lime, 25);
  rect(s, 12.192, 10.092, 6.521, 3.8, C.green, 65);
  crumbs(s, 'Chapter 05', 'Competitive Advantages', CRUMB.wide);
  rect(s, 12.192, 2.764, 13.189, 7.189, C.pale);
  txt(s, 'Operations', 12.608, 3.216, 10.971, 2.125, S.display);
  txt(s, 'Etiam dignissim diam quis enim lobortis scelerisque fermentum. Non quam lacus suspendisse faucibus interdum posuere. Et sollicitudin ac orci phasellus egestas tellus rutrum. Magnis dis parturient montes.',
    12.608, 7.73, 7.99, 1.772, S.small);
  txt(s, 'Stay Healthy with Green Eats', 12.608, 10.566, 3.775, 1.194,
    { fontFace: F.reg, fontSize: 32, bold: true, color: C.white });
  txt(s, 'Etiam dignissim diam quis enim lobortis scelerisque fermentum. Non quam lacus suspendisse faucibus interdum posuere.',
    12.608, 12.473, 5.688, 0.944, { fontFace: F.reg, fontSize: 16, color: C.white });
  txt(s, 'Etiam dignissim diam.', 19.1, 11.219, 5.688, 0.389, { fontFace: F.reg, fontSize: 16, color: C.white });
  txt(s, '12 000', 19.1, 11.758, 5.942, 2.125, S.display, { color: C.white, align: 'center' });
}

// 14 — Management team
function slide14(pres) {
  const s = pres.addSlide();
  crumbs(s, 'Chapter 05', 'Management Team', CRUMB.team);
  txt(s, 'PLACEHOLDER',
    1.286, 3.488, 8.947, 2.569, { fontFace: F.reg, fontSize: 48, color: C.slate });

  const team = [
    { x: 1.286, name: 'Fredrik P Svensson', role: 'Chief Executive Officer', chip: C.silver },
    { x: 7.301, name: 'Patrick Petit ', role: 'Business Developer', chip: C.lime },
    { x: 13.316, name: 'Tobias Lindelof', role: 'Chief Financial Officer', chip: C.silver },
    { x: 19.331, name: 'Yasin Roberts', role: 'Executive Chef', chip: C.silver },
    { x: 25.346, name: 'Fredrik P Svensson', role: 'Chief Executive Officer', chip: C.silver }
  ];
  team.forEach(function (m) {
    photo(s, m.x, 6.99, 5.613, 3.889);
    rect(s, m.x, 10.879, 5.613, 3.012, C.mist);
    rect(s, m.x, 12.847, 5.613, 1.044, C.pale);
    rect(s, m.x + 4.571, 12.847, 1.042, 1.044, m.chip);
    txt(s, m.name, m.x + 0.386, 11.296, 4.354, 0.472,
      { fontFace: F.reg, fontSize: 22, bold: true, color: C.body, lineSpacingMultiple: 1.2 });
    txt(s, m.role, m.x + 0.386, 11.768, 4.354, 0.389, S.micro);
    txt(s, 'View Details', m.x + 0.386, 13.175, 2.392, 0.389, S.micro);
    arrowGlyph(s, m.x + 4.961, 13.196, 0.263, 0.346, 90, C.black);
  });
}

// 15 — Financial projections
function slide15(pres) {
  const s = pres.addSlide();
  crumbs(s, 'Chapter 05', 'Financial Projections', CRUMB.fin);
  txt(s, 'Projections', 1.286, 2.033, 18.03, 2.125, S.display);

  const tabs = [
    { x: 20.731, label: 'Week', on: false },
    { x: 22.459, label: 'Month', on: false },
    { x: 24.188, label: 'Year', on: true }
  ];
  tabs.forEach(function (t) {
    txt(s, t.label, t.x, 2.859, 1.729, 0.472, t.on
      ? { fontFace: F.reg, fontSize: 22, bold: true, color: C.black, align: 'center', lineSpacingMultiple: 1.2, underline: { style: 'sng' } }
      : { fontFace: F.light, fontSize: 22, color: C.grey, align: 'center', lineSpacingMultiple: 1.2 });
  });

  txt(s, 'Understanding the target market and consumer trends is crucial to the success of GreenEats. Here is an overview of the market landscape and the demand for sustainable and healthy food options.',
    1.286, 4.654, 11.399, 2.357, S.lead, { valign: 'top', margin: 7.2 });
  txt(s, '56.90%', 17.288, 4.722, 6.584, 2.132, S.display);
  s.addShape('triangle', { x: 23.872, y: 5.092, w: 0.346, h: 0.3,
    fill: { color: C.green }, line: { color: C.silver, width: 0.75 } });
  txt(s, 'Project for next 4 years', 23.872, 5.865, 1.509, 0.651, S.micro);
  photo(s, 1.286, 7.419, 24.095, 3.889);
  arrowGlyph(s, 30.307, 13.196, 0.263, 0.346, 90, C.black);

  const note = 'Etiam dignissim diam quis enim lobortis scelerisque fermentum. Non quam lacus suspendisse faucibus interdum posuere.';
  const cards = [
    { x: 1.284, textX: 1.561, noteH: 0.92, stampH: 0.382 },
    { x: 7.44, textX: 7.717, noteH: 0.92, stampH: 0.382 },
    { x: 13.596, textX: 13.873, noteH: 0.92, stampH: 0.382 },
    { x: 19.752, textX: 20.029, noteH: 0.945, stampH: 0.389 }
  ];
  cards.forEach(function (c) {
    rect(s, c.x, 11.716, 5.629, 2.175, C.mist);
    txt(s, note, c.textX, 11.991, 5.073, c.noteH, S.micro);
    txt(s, '10 Aug, 6:11 AM', c.textX, 13.228, 5.073, c.stampH, S.micro);
  });
}

// 16 — Funding requirements
function slide16(pres) {
  const s = pres.addSlide();
  crumbs(s, 'Chapter 05', 'Funding Requirements', CRUMB.fin);
  txt(s, 'Total Funding Needed :', 1.286, 4.726, 8.786, 0.653, S.lead);
  txt(s, '$2,500,000', 1.286, 5.379, 10.695, 2.125, S.display);
  txt(s, 'Etiam dignissim diam quis enim lobortis scelerisque fermentum. Non quam lacus suspendisse faucibus interdum posuere.',
    1.284, 8.537, 10.402, 1.736, S.lead);
  arrowGlyph(s, 30.307, 13.196, 0.263, 0.346, 90, C.black);

  const marketingCopy = 'Marketing campaigns and strategies to increase brand awareness, attract new customers, and build customer loyalty';
  const hiringCopy = 'Recruitment and hiring of skilled professionals across various departments, such as operations, marketing, and customer service';
  const cards = [
    { x: 12.38, y: 1.108, n: '01', label: 'Infrastructure', labelY: 1.691, labelH: 0.472, pct: '40%',
      body: 'Investment in operational infrastructure, including kitchen facilities, equipment, and technology systems.' },
    { x: 16.827, y: 1.108, n: '02', label: 'Marketing', labelY: 1.691, labelH: 0.472, pct: '25%', body: marketingCopy },
    { x: 21.211, y: 1.108, n: '03', label: 'Talent Acquisition', labelY: 1.691, labelH: 0.472, pct: '15%', body: hiringCopy },
    { x: 12.38, y: 8.198, n: '04', label: 'Research and Development', labelY: 8.564, labelH: 0.906, pct: '10%',
      body: 'Enhance menu offerings, explore new sustainable sourcing partnerships, and improve operational efficiency.' },
    { x: 16.827, y: 8.198, n: '05', label: 'Branding', labelY: 8.78, labelH: 0.472, pct: '32%', body: marketingCopy },
    { x: 21.211, y: 8.198, n: '06', label: 'In-House Training', labelY: 8.78, labelH: 0.472, pct: '5%', body: hiringCopy }
  ];
  cards.forEach(function (c) {
    rect(s, c.x, c.y, 4.169, 6.812, C.lime);
    txt(s, c.n, c.x + 0.417, c.y + 0.417, 0.804, 0.804,
      { shape: 'ellipse', fill: { color: C.green }, fontFace: F.reg, fontSize: 22, bold: true,
        color: C.white, align: 'center', margin: SHAPE_INSET, lineSpacingMultiple: 1.2 });
    txt(s, c.label, c.x + 1.448, c.labelY, 2.529, c.labelH, S.small);
    txt(s, c.pct, c.x + 0.417, c.y + 1.669, 3.336, 1.181, { fontFace: F.semi, fontSize: 64, color: C.slate });
    txt(s, c.body, c.x + 0.417, c.y + 3.406, 3.336, 2.639, S.small);
  });
}

// 17 — Timeline (gantt)
function slide17(pres) {
  const s = pres.addSlide();
  crumbs(s, 'Chapter 05', 'Timeline', CRUMB.narrow);
  txt(s, 'Timeline', 1.286, 3.321, 9.007, 2.125, S.display);
  txt(s, 'We provide a solution for individuals who seek convenient, nutritious meals without compromising on their commitment to sustainability.',
    1.284, 6.647, 6.903, 2.819, S.lead);
  arrowGlyph(s, 30.307, 13.196, 0.263, 0.346, 90, C.black);

  [11.96, 13.875, 15.79, 17.706, 19.621, 21.536, 23.451, 25.367].forEach(function (x) {
    rule(s, x, 3.321, 0, 8.359, C.silver);
  });
  [[12.229, 1.36], [14.13, 1.406], [16.077, 1.36], [17.983, 1.36],
    [19.875, 1.406], [21.814, 1.36], [23.706, 1.406]].forEach(function (t) {
    txt(s, 'X 01', t[0], 3.557, t[1], 0.472,
      { fontFace: F.light, fontSize: 22, color: C.body, align: 'center', lineSpacingMultiple: 1.2 });
  });

  const rows = [
    { label: 'Marketing', labelY: 4.81, labelH: 0.472, bars: [[12.587, 4.692, 3.821, C.silver]] },
    { label: 'Recruitment', labelY: 6.122, labelH: 0.472, bars: [[14.833, 5.975, 4.771, C.silver], [20.579, 5.975, 4.771, C.silver]] },
    { label: 'Branding', labelY: 7.406, labelH: 0.472, bars: [[11.974, 7.259, 8.605, C.lime]] },
    { label: 'Business Development', labelY: 8.472, labelH: 0.906, bars: [[19.607, 8.542, 5.746, C.silver]] },
    { label: 'Go-to Market', labelY: 9.973, labelH: 0.472, bars: [[17.692, 9.826, 4.771, C.silver]] }
  ];
  rows.forEach(function (r) {
    r.bars.forEach(function (b) { rect(s, b[0], b[1], b[2], 0.765, b[3]); });
    txt(s, r.label, 9.453, r.labelY, 2.493, r.labelH, S.small);
  });
}

// 18 — Investor highlights
function slide18(pres) {
  const s = pres.addSlide();
  crumbs(s, 'Chapter 05', 'Funding', CRUMB.narrow);
  txt(s, 'Investor Highlights', 1.286, 2.953, 18.019, 2.132, S.display);
  arrowGlyph(s, 30.307, 13.196, 0.263, 0.346, 90, C.black);

  const longCopy = 'Understanding the target market and consumer trends is crucial to the success of GreenEats. Here is an overview of the market landscape and the demand for sustainable and healthy food options.';
  const shortCopy = 'Understanding the target market and consumer trends is crucial to the success of GreenEats. Here is an overview of the market landscape and the demand.';
  const cards = [
    { x: 0, y: 7.96, w: 9.051, fill: null, textX: 1.286, textW: 6.735, textH: 3.903, stamp: C.grey, copy: longCopy },
    { x: 9.051, y: 7.96, w: 8.167, fill: C.lime, textX: 9.482, textW: 7.306, textH: 2.819, stamp: C.black, copy: shortCopy },
    { x: 17.197, y: 7.946, w: 8.167, fill: null, textX: 17.615, textW: 7.306, textH: 2.819, stamp: C.grey, copy: shortCopy }
  ];
  cards.forEach(function (c) {
    if (c.fill) rect(s, c.x, c.y, c.w, 5.932, c.fill);
    else frame(s, c.x, c.y, c.w, 5.932, C.silver);
    txt(s, '10 Aug, 6:11 AM', c.textX, 8.362, 3.24, 0.389, S.micro, { color: c.stamp });
    txt(s, c.copy, c.textX, 9.335, c.textW, c.textH, S.lead);
  });
}

// 19 — Marketing quadrants
function slide19(pres) {
  const s = pres.addSlide();
  const colX = [1.3, 9.037, 16.775, 24.513];
  [5.543, 9.457, 13.372].forEach(function (y) {
    colX.forEach(function (x) {
      // the top-left cell is a solid lime block instead of an outline
      if (x === 1.3 && y === 5.543) return;
      frame(s, x, y, 7.738, 3.915, C.silver);
    });
  });
  rect(s, 1.3, 5.543, 7.738, 3.915, C.lime);
  crumbs(s, 'Chapter 05', 'Timeline', CRUMB.narrow);
  txt(s, 'Marketing', 1.286, 2.59, 15.254, 2.125, S.display);

  const cells = [
    { titleX: 1.703, titleY: 5.963, titleW: 5.378, title: 'Social Media Strategies',
      bodyX: 1.703, bodyY: 7.278, bodyH: 1.772,
      body: 'We develop compelling social media content that showcases our sustainable food offerings, promotes healthy lifestyles, and educates our audience about the importance of sustainability' },
    { titleX: 17.233, titleY: 5.963, titleW: 6.013, title: 'Online Advertising',
      bodyX: 17.218, bodyY: 7.711, bodyH: 1.339,
      body: 'We leverage data-driven advertising platforms to reach our target market with precision and deliver personalized messages.' },
    { titleX: 9.468, titleY: 9.884, titleW: 6.013, title: 'Public Relations and Media',
      bodyX: 9.468, bodyY: 11.668, bodyH: 1.339,
      body: 'We regularly issue press releases to announce key milestones, partnerships, and sustainability initiatives. Lorem ipsum.' },
    { titleX: 17.218, titleY: 9.884, titleW: 6.013, title: 'Influencer Partnerships',
      bodyX: 17.218, bodyY: 11.668, bodyH: 1.339,
      body: 'Collaborating with relevant influencers who align with our brand values allows us to expand our reach and tap into their engaged audience.' }
  ];
  cells.forEach(function (c) {
    txt(s, c.title, c.titleX, c.titleY, c.titleW, 0.653, { fontFace: F.light, fontSize: 32, color: C.black });
    txt(s, c.body, c.bodyX, c.bodyY, 6.773, c.bodyH, S.small);
  });
}

// 20 — Contact
function slide20(pres) {
  const s = pres.addSlide();
  rect(s, 1.3, 5.543, 24.067, 8.349, C.lime);
  photo(s, 17.326, 5.529, 8.055, 8.363);
  crumbs(s, 'Chapter 05', 'Timeline', CRUMB.narrow);
  txt(s, 'Green for Future', 1.286, 2.59, 15.254, 2.125, S.display);
  txt(s, 'For inquiries, collaborations, or to learn more about GreenEats, please find our contact information below',
    1.967, 6.494, 8.674, 1.736, S.lead);

  const contacts = [
    { y: 10.237, label: 'Phone', value: ': +7 (8202) 53 0900' },
    { y: 11.899, label: 'Fax', value: ': +7 (8202) 53 0900' },
    { y: 12.551, label: 'Email', value: ': info@greeneats.com' }
  ];
  contacts.forEach(function (c) {
    txt(s, c.label, 1.967, c.y, 1.537, 0.653, { fontFace: F.reg, fontSize: 32, bold: true, color: C.black });
    txt(s, c.value, 3.504, c.y, 6.013, 0.653, { fontFace: F.light, fontSize: 32, color: C.black });
  });
  txt(s, 'Lorem ipsum dolor sitem amet.', 1.967, 10.889, 6.013, 0.472, S.small);

  txt(s, 'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat.',
    18.586, 9.984, 6.013, 0.906, S.small, { color: C.white });
  const socials = [
    { y: 11.246, label: 'Instagram', value: ': @greeneats' },
    { y: 11.899, label: 'Facebook', value: ': greeneats' },
    { y: 12.551, label: 'Twitter', value: ': @greeneats' }
  ];
  socials.forEach(function (c) {
    txt(s, c.label, 18.586, c.y, 2.62, 0.653, { fontFace: F.reg, fontSize: 32, bold: true, color: C.white });
    txt(s, c.value, 21.206, c.y, 4.578, 0.653, { fontFace: F.light, fontSize: 32, color: C.white });
  });
}

// 21 & 22 — icon library sheets (outline discs with simple inner pictograms)
function slide21(pres) {
  const s = pres.addSlide();
  const even = [2.670, 4.891, 7.113, 9.334, 11.555, 13.776, 15.998, 18.219, 20.441, 22.661];
  iconGrid(s, 1.315, [
    { y: 2.164, xs: even,
      marks: ['box', 'bars', 'ring', 'cross', 'list', 'dot', 'minus', 'tri', 'dot', 'ring'] },
    { y: 4.550, xs: [2.658, 4.884, 7.110, 9.336, 11.562, 13.789, 16.015, 18.241, 20.467, 22.694],
      marks: ['bars', 'ring', 'dot', 'box', 'boxFilled', 'box', 'ring', 'right', 'bars', 'down'] },
    { y: 6.848, xs: [2.670, 4.916, 7.005, 9.148, 11.404, 13.751, 16.059, 18.445, 20.514, 22.661],
      marks: ['down', 'tri', 'cross', 'ring', 'down', 'cross', 'minus', 'cross', 'boxFilled', 'bars'] },
    { y: 9.135, xs: even,
      marks: ['box', 'tri', 'upRight', 'plus', 'up', 'bars', 'ring', 'ring', 'right', 'upRight'] },
    { y: 11.521, xs: even,
      marks: ['left', 'left', 'upLeft', 'down', 'downLeft', 'bars', 'minus', 'ring', 'plus', 'plus'] }
  ]);
}

function slide22(pres) {
  const s = pres.addSlide();
  const even = [1.470, 3.966, 6.462, 8.958, 11.454, 13.950, 16.446, 18.942, 21.438, 23.934];
  iconGrid(s, 1.281, [
    { y: 1.722, xs: [1.470, 3.969, 6.467, 8.966, 11.465, 13.964, 16.463, 18.962, 21.461, 23.934],
      marks: ['bars', 'dot', 'right', 'ring', 'up', 'box', 'plus', 'tri', 'ring', 'minus'] },
    { y: 4.318, xs: even,
      marks: ['minus', 'plus', 'ring', 'dot', 'box', 'boxFilled', 'dot', 'dot', 'ring', 'tri'] },
    { y: 6.780, xs: even,
      marks: ['bars', 'ring', 'minus', 'ring', 'dot', 'plus', 'up', 'right', 'dot', 'tri'] },
    { y: 9.418, xs: even,
      marks: ['tri', 'box', 'upRight', 'upRight', 'tri', 'bars', 'bars', 'bars', 'list', 'dot'] },
    { y: 11.997, xs: [1.451, 3.949, 6.447, 8.945, 11.444, 13.942, 16.440, 18.938, 21.436, 23.934],
      marks: ['box', 'box', 'list', 'bars', 'box', 'dot', 'dot', 'dot', 'bars', 'ring'] }
  ]);
}

/* --------------------------------------------------------------------- main */
function build() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'GREENEATS', width: 26.667, height: 15 });
  pres.layout = 'GREENEATS';
  pres.author = 'GreenEats';
  pres.title = 'GreenEats';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
    slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
    slide17, slide18, slide19, slide20, slide21, slide22].forEach(function (fn) { fn(pres); });

  return pres.writeFile({
    fileName: path.join(__dirname, '09531d82-da3a-4979-8692-026d9999aafa_grok_final.pptx')
  });
}

build().then(function (f) { console.log('wrote ' + f); }).catch(function (e) { console.error(e); process.exit(1); });
