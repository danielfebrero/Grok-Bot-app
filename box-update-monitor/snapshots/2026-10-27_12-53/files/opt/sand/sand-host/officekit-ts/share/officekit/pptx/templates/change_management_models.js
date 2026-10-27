'use strict';
/*
 * "Change Management Models" - 27 slide deck rebuilt with pptxgenjs.
 * Slide size 10 x 5.625in (16:9). Colours come from the deck theme "CMM".
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ── palette / type ──────────────────────────────────────────────────────────
const BLACK = '050708';
const LIGHT = 'E9E9E9';
const GREY = '5E5E5E';
const LIME = 'BCE81F';
const SOFT = 'EAEAEA';
const LIME2 = 'C6E750';
const FONT = 'Rajdhani Medium';
const SW = 10, SH = 5.625;

// ── generic drawing helpers ─────────────────────────────────────────────────
function txt(s, content, o) {
  s.addText(content, Object.assign({ fontFace: FONT, margin: 1.5, valign: 'top', align: 'left', color: BLACK, lineSpacingMultiple: 1 }, o));
}
function rect(s, o) { s.addShape('rect', o); }
function ell(s, o) { s.addShape('ellipse', o); }
function ln(s, o) { s.addShape('line', o); }
// OOXML "adj" values are 1/100000 of the shorter side, pptxgenjs wants inches.
function rrect(s, o) {
  const adj = o.adj === undefined ? 16667 : o.adj;
  s.addShape('roundRect', Object.assign({}, o, { rectRadius: (adj / 100000) * Math.min(o.w, o.h) }));
}
// homePlate / chevron take a single adjust = point (and notch) depth in inches
function tag(s, shape, o) { s.addShape(shape, Object.assign({}, o, { rectRadius: o.point })); }
// custGeom from rings of [0..1, 0..1] normalised points
function poly(s, o, rings) {
  const pts = [];
  rings.forEach(function (ring) {
    ring.forEach(function (p, i) { pts.push({ x: +(o.w * p[0]).toFixed(4), y: +(o.h * p[1]).toFixed(4), moveTo: i === 0 }); });
    pts.push({ close: true });
  });
  s.addShape('custGeom', Object.assign({}, o, { points: pts }));
}

// ── decorative motifs ───────────────────────────────────────────────────────
// clip a polygon to the vertical strip 0 <= x <= w (Sutherland-Hodgman)
function clipX(pts, w) {
  [true, false].forEach(function (left) {
    const inside = function (p) { return left ? p[0] >= 0 : p[0] <= w; };
    const lim = left ? 0 : w, out = [];
    for (let i = 0; i < pts.length; i++) {
      const a = pts[i], b = pts[(i + 1) % pts.length];
      if (inside(a)) out.push(a);
      if (inside(a) !== inside(b)) {
        const t = (lim - a[0]) / (b[0] - a[0]);
        out.push([lim, a[1] + t * (b[1] - a[1])]);
      }
    }
    pts = out;
  });
  return pts;
}
// 45 degree "/" hatch filling a rectangle - used as the deck's texture panels
function stripes(s, o) {
  const period = 0.3, thick = 0.147, rings = [];
  for (let c = 0; c < o.w + o.h; c += period) {
    const band = clipX([[c, 0], [c + thick, 0], [c + thick - o.h, o.h], [c - o.h, o.h]], o.w);
    if (band.length > 2) rings.push(band.map(function (p) { return [p[0] / o.w, p[1] / o.h]; }));
  }
  poly(s, o, rings);
}
// thin "down" arrow glyph (stem + open head) used as accent all over the deck
const ARROW = [[0.955, 0.603], [0.532, 0.912], [0.532, 0], [0.468, 0], [0.468, 0.912], [0.045, 0.603], [0, 0.636], [0.5, 1], [1, 0.636]];
function arrow(s, o) { poly(s, o, [ARROW]); }
// tick mark
const TICK = [[0.943, 0], [0.360, 0.837], [0.057, 0.401], [0, 0.483], [0.360, 1], [1, 0.081]];
function tick(s, o) { poly(s, o, [TICK]); }
// long thin arrow with an open V head (process timeline on slides 7 & 8)
const THIN = [[0.859, 0], [0.846, 0.045], [0.966, 0.468], [0, 0.468], [0, 0.532], [0.966, 0.532], [0.846, 0.955], [0.859, 1], [1, 0.5]];
function thinArrow(s, o) { poly(s, o, [THIN]); }
// badge with a tick (contact / testimonial marker)
function shieldTick(s, o) {
  poly(s, { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.color || LIME } },
    [[[0, 0], [1, 0], [1, 0.68], [0.5, 1], [0, 0.68]]]);
  tick(s, { x: o.x + o.w * 0.18, y: o.y + o.h * 0.22, w: o.w * 0.64, h: o.h * 0.34, fill: { color: o.tickColor || BLACK } });
}
function quote(s, o) {
  txt(s, '❞', { x: o.x - 0.10, y: o.y - 0.10, w: o.w + 0.20, h: o.h + 0.20, fontSize: 14, bold: true, color: o.color, align: 'center', valign: 'middle' });
}

// ── page furniture (footer + slide number), keyed by background ─────────────
const CHROME = {
  light: { arrow: GREY, label: GREY, num: BLACK, bg: LIGHT },
  dark: { arrow: LIME, label: GREY, num: LIGHT, bg: BLACK },
  grey: { arrow: LIME, label: LIGHT, num: LIGHT, bg: GREY },
  lime: { arrow: GREY, label: GREY, num: BLACK, bg: LIME }
};
function page(pptx, kind) {
  const c = CHROME[kind];
  const s = pptx.addSlide();
  s.background = { color: c.bg };
  arrow(s, { x: 0.668, y: 5.12, w: 0.087, h: 0.12, rotate: -45, fill: { color: c.arrow } });
  txt(s, 'Change Management Models', { x: 0.797, y: 5.086, w: 1.827, h: 0.198, fontSize: 9, color: c.label });
  s.slideNumber = { x: 9.018, y: 5.086, w: 0.375, h: 0.198, fontFace: FONT, fontSize: 9, color: c.num, align: 'center' };
  return s;
}
// Photo frames: the source deck ships them as *empty* picture placeholders, which
// render as nothing. Keep the frames in the layout but leave them unpainted.
function photo(s, o) {
  rrect(s, { x: o.x, y: o.y, w: o.w, h: o.h, adj: o.adj, fill: { color: GREY, transparency: 100 } });
}

// ── slide 1 · cover ─────────────────────────────────────────────────────────
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: LIME };
  stripes(s, { x: 5.537, y: 0, w: 4.463, h: 5.625, fill: { color: SOFT, transparency: 70 } });
  txt(s, 'Change Management Models', { x: 1.212, y: 1.989, w: 5.908, h: 1.683, fontSize: 49, valign: 'middle' });
  ln(s, { x: 0.682, y: 0, w: 0, h: 5.625, line: { color: BLACK, width: 1.5 } });
  rect(s, { x: 0, y: 0, w: 0.691, h: 0.664, fill: { color: BLACK } });
  // hamburger glyph inside the black corner block
  poly(s, { x: 0.171, y: 0.167, w: 0.351, h: 0.311, fill: { color: LIME } },
    [[[0, 0], [1, 0], [1, 0.099], [0, 0.099]], [[0, 0.451], [1, 0.451], [1, 0.549], [0, 0.549]], [[0, 0.901], [1, 0.901], [1, 1], [0, 1]]]);
  arrow(s, { x: 1.294, y: 3.776, w: 0.491, h: 0.674, rotate: -45, fill: { color: BLACK } });
  // social icons down the left rail
  ['f', '♪', '✦'].forEach(function (glyph, i) {
    txt(s, glyph, { x: 0.19, y: 3.661 + i * 0.5, w: 0.31, h: 0.2, fontSize: 11, bold: true, align: 'center', valign: 'middle' });
  });
  rrect(s, { x: 0.235, y: 5.16, w: 0.22, h: 0.22, adj: 25000, fill: { color: LIME, transparency: 100 }, line: { color: BLACK, width: 1 } });
  ell(s, { x: 0.30, y: 5.225, w: 0.09, h: 0.09, fill: { color: BLACK } });
  txt(s, '27 Slides With Value Convention', { x: 1.238, y: 0.283, w: 2.27, h: 0.24, fontSize: 11, valign: 'middle' });
  [['Home', 4.481], ['About us', 5.828], ['Company', 7.175], ['Contact', 8.522]].forEach(function (n) {
    txt(s, n[0], { x: n[1], y: 0.283, w: 1.038, h: 0.24, fontSize: 11, valign: 'middle' });
  });
  rrect(s, { x: 4.282, y: 0.34, w: 0.127, h: 0.127, adj: 25312, fill: { color: LIGHT } });
}

// ── slide 2 · intro ─────────────────────────────────────────────────────────
function slide02(pptx) {
  const s = page(pptx, 'dark');
  txt(s, 'Understanding Change Management Models', { x: 0.612, y: 1.605, w: 3.661, h: 2.566, fontSize: 38, color: LIGHT });
  txt(s, 'Change Management Models are strategic frameworks designed to facilitate and manage organizational change. ' +
    'They provide a structured approach to transition individuals, teams, or organizations from a current state to a desired future state. ' +
    'These models offer guidance through different stages of change, aiding in managing resistance, engagement, and adaptation to ensure ' +
    'effective transformation. Crucial in times of transition, they can maximize success and minimize disruption.',
    { x: 4.989, y: 2.694, w: 3.873, h: 1.355, fontSize: 10, color: GREY });
  // lime toggle pill with a black knob + tick
  rrect(s, { x: 0.677, y: 1.01, w: 0.957, h: 0.451, adj: 50000, fill: { color: LIME } });
  ell(s, { x: 0.714, y: 1.047, w: 0.377, h: 0.377, fill: { color: BLACK } });
  tick(s, { x: 0.785, y: 1.16, w: 0.223, h: 0.156, fill: { color: LIME } });
  stripes(s, { x: 9.537, y: 0, w: 0.463, h: 5.625, fill: { color: GREY, transparency: 70 } });
}

// ── slides 3 & 4 · team ─────────────────────────────────────────────────────
function person(s, o) {
  txt(s, [{ text: o.first, options: { breakLine: true } }, { text: o.last }],
    { x: o.x + 0.005, y: o.y, w: 1.869, h: 0.547, fontSize: 15 });
  ln(s, { x: o.x + 0.023, y: o.y + 0.642, w: 1.791, h: 0, line: { color: BLACK, width: 1.5 } });
  txt(s, o.role, { x: o.x, y: o.y + 0.754, w: 1.869, h: 0.214, fontSize: 10, color: GREY });
}
const TEAM = [
  { first: 'James', last: 'Allen', role: 'Senior Strategy Consultant' },
  { first: 'Sophia', last: 'Martinez', role: 'Business Transformation Lead' },
  { first: 'Richard', last: 'Thompson', role: 'Chief Innovation Consultant' },
  { first: 'Chloe', last: 'Sullivan', role: 'Director of Market Insights' }
];
function slide03(pptx) {
  const s = page(pptx, 'light');
  [[0.668, 12877], [3.006, 12294], [5.343, 12877]].forEach(function (p) {
    photo(s, { x: p[0], y: 0.615, w: 2.042, h: 3.032, adj: p[1] });
  });
  TEAM.slice(0, 3).forEach(function (t, i) { person(s, Object.assign({ x: 0.877 + i * 2.284, y: 3.828 }, t)); });
  txt(s, 'Our Team', { x: 6.314, y: 1.907, w: 3.671, h: 0.989, fontSize: 56, rotate: 90 });
  txt(s, '01', { x: 8.307, y: 0.406, w: 1.104, h: 0.989, fontSize: 56, color: LIME, align: 'right' });
  stripes(s, { x: 8.74, y: 0, w: 0.639, h: 0.469, fill: { color: GREY, transparency: 80 } });
}
function slide04(pptx) {
  const s = page(pptx, 'light');
  [[0.646, 0.716, 13336], [2.892, 0.716, 12780], [5.137, 1.832, 13058], [7.382, 1.832, 12503]].forEach(function (p) {
    photo(s, { x: p[0], y: p[1], w: 1.958, h: 2.052, adj: p[2] });
  });
  TEAM.forEach(function (t, i) { person(s, Object.assign({ x: 0.789 + i * 2.243, y: i < 2 ? 2.922 : 4.037 }, t)); });
  txt(s, 'Our Team', { x: 5.246, y: 0.477, w: 3.671, h: 0.989, fontSize: 56 });
  txt(s, '01', { x: 8.994, y: 0.478, w: 1.104, h: 0.989, fontSize: 56, color: LIME, align: 'right' });
}

// ── slide 5 · product name ──────────────────────────────────────────────────
const PRODUCT_FEATURES = [
  ['Solution-oriented: ', 'The product aims to solve specific business problems, thereby improving efficiency and effectiveness.'],
  ['Scalable: ', 'The product can adapt and grow with the business, ensuring long-term usability.'],
  ['Customizable: ', 'It can be tailored to the unique needs and objectives of each business.'],
  ['Data-driven: ', 'Utilizes data analysis to provide insights and recommendations, assisting in informed decision-making.'],
  ['User-friendly: ', 'Designed for ease-of-use, making it accessible to different user levels within an organization.'],
  ['Integration Capable: ', 'Able to integrate with other systems, streamlining business operations.']
];
function slide05(pptx) {
  const s = page(pptx, 'dark');
  photo(s, { x: 0.656, y: 0.812, w: 2.96, h: 3.929, adj: 9114 });
  txt(s, 'Product Name', { x: 4.215, y: 0.759, w: 3.671, h: 0.673, fontSize: 38, color: LIGHT });
  PRODUCT_FEATURES.forEach(function (f, i) {
    const y = 1.523 + i * 0.5547;
    txt(s, [{ text: f[0], options: { color: LIGHT } }, { text: f[1], options: { color: GREY } }],
      { x: 4.228, y: y, w: 3.873, h: 0.37, fontSize: 10 });
    tick(s, { x: 8.629, y: y + 0.115, w: 0.223, h: 0.156, fill: { color: LIME } });
    if (i < 5) ln(s, { x: 4.244, y: y + 0.472, w: 4.825, h: 0, line: { color: GREY, width: 1.5, transparency: 70 } });
  });
  stripes(s, { x: 9.537, y: 0, w: 0.463, h: 5.625, fill: { color: GREY, transparency: 70 } });
}

// ── slide 6 · products ──────────────────────────────────────────────────────
const PRODUCT_BULLETS = [
  ['Adaptable: ', 'Able to adjust to different business environments and needs.'],
  ['Efficient: ', 'Streamlines business processes for improved productivity.'],
  ['Insightful: ', 'Provides valuable data and insights for decision-making.'],
  ['Reliable: ', 'Dependable under various operational conditions.']
];
function slide06(pptx) {
  const s = page(pptx, 'light');
  txt(s, 'Our Products', { x: 0.612, y: 0.485, w: 3.661, h: 0.673, fontSize: 38 });
  txt(s, '03', { x: -0.107, y: 0.485, w: 0.779, h: 0.673, fontSize: 38, color: LIME });
  ['/A', '/B', '/C'].forEach(function (letter, i) {
    const x = 0.664 + i * 2.238;
    rrect(s, { x: x, y: 1.313, w: 1.94, h: 3.456, adj: 13425, fill: { color: BLACK } });
    txt(s, letter, { x: x + 0.209, y: 1.507, w: 0.609, h: 0.295, fontSize: 15, color: LIME });
    txt(s, 'Product', { x: x + 0.209, y: 1.756, w: 1.398, h: 0.421, fontSize: 23, color: LIGHT });
    const body = [];
    PRODUCT_BULLETS.forEach(function (b, j) {
      if (j) body.push({ text: '', options: { breakLine: true } });
      body.push({ text: b[0], options: { color: LIGHT } });
      body.push({ text: b[1], options: { color: GREY, breakLine: true } });
    });
    txt(s, body, { x: x + 0.214, y: 2.195, w: 1.54, h: 1.683, fontSize: 8 });
    rrect(s, { x: x + 0.219, y: 4.143, w: 1.502, h: 0.408, adj: 27312, fill: { color: LIME } });
    txt(s, 'More Information', { x: x + 0.291, y: 4.249, w: 1.348, h: 0.198, fontSize: 9, align: 'center' });
  });
  rrect(s, { x: 7.377, y: 1.313, w: 1.94, h: 3.456, adj: 13425, fill: { color: LIME } });
  stripes(s, { x: 7.437, y: 1.353, w: 1.82, h: 3.376, fill: { color: LIGHT, transparency: 70 } });
  txt(s, 'Need anything else?', { x: 7.679, y: 2.704, w: 1.502, h: 1.178, fontSize: 23 });
  arrow(s, { x: 7.782, y: 4.097, w: 0.349, h: 0.478, rotate: -90, fill: { color: BLACK } });
}

// ── slides 7 & 8 · step-by-step guide ───────────────────────────────────────
const GUIDE_STEPS = [
  ['01', 'Identify the Need for Change', 'Understand the current challenges and why change is necessary.'],
  ['02', 'Define Desired Outcomes', 'Clearly outline the objectives and goals of the change.'],
  ['03', 'Assess Organizational Culture', "Consider the company's culture and how it may affect change implementation."],
  ['04', 'Evaluate Resources and Capabilities', 'Understand the resources and skills available to manage the change.'],
  ['05', 'Understand Employee Dynamics', 'Consider the attitudes, fears, and expectations of the employees towards the change.'],
  ['06', 'Review Past Change Experiences', 'Learn from previous change initiatives, their successes, and failures.'],
  ['07', 'Align with Strategic Vision', "Ensure the chosen model supports the company's long-term vision and strategy."]
];
// one column of the timeline: dashes, arrow, number, heading, body
function guideStep(s, o) {
  const dark = o.active;
  thinArrow(s, { x: o.x + 0.606, y: 2.6, w: 1.236, h: 0.349, fill: { color: dark ? BLACK : GREY, transparency: dark ? 0 : 50 } });
  [[0, 0.1], [0.21, 0.082], [0.403, 0.082]].forEach(function (d) {
    rect(s, { x: o.x + 0.01 + d[0], y: 2.763, w: d[1], h: 0.022, fill: { color: dark ? BLACK : GREY, transparency: dark ? 0 : 20 } });
  });
  txt(s, o.step[0], { x: o.x - 0.037, y: 2.115, w: 0.621, h: 0.547, fontSize: 30, color: LIME });
  txt(s, o.step[1], { x: o.x - 0.001, y: 3.018, w: 1.729, h: 0.799, fontSize: 15, valign: 'bottom' });
  txt(s, o.step[2], { x: o.x, y: 3.917, w: 1.837, h: 0.534, fontSize: 10, color: GREY });
}
function slide07(pptx) {
  const s = page(pptx, 'light');
  txt(s, 'Selecting the Ideal Change Management Model:\nA Step-by-Step Guide',
    { x: 0.612, y: 0.714, w: 6.252, h: 0.799, fontSize: 23 });
  GUIDE_STEPS.slice(0, 4).forEach(function (st, i) {
    guideStep(s, { x: 0.656 + i * 2.1155, step: st, active: i === 0 });
  });
  thinArrow(s, { x: 9.772, y: 2.6, w: 1.236, h: 0.349, fill: { color: GREY, transparency: 50 } });
}
function slide08(pptx) {
  const s = page(pptx, 'light');
  GUIDE_STEPS.slice(4).forEach(function (st, i) {
    guideStep(s, { x: 0.656 + i * 2.1155, step: st, active: i === 2 });
  });
  thinArrow(s, { x: -1.462, y: 2.6, w: 1.833, h: 0.349, fill: { color: GREY, transparency: 50 } });
  txt(s, '01', { x: -1.508, y: 2.115, w: 0.621, h: 0.573, fontSize: 30, color: LIME2 });
  txt(s, 'Choosing the right Change Management Model is critical to the success of your business transformation. ' +
    "It involves understanding your organization's specific needs, culture, and resources, and aligning them with your strategic vision. " +
    'This comprehensive approach ensures a smoother transition, enhanced employee engagement, and effective management of change.',
    { x: 7.515, y: 2.46, w: 1.942, h: 2.011, fontSize: 10, color: GREY });
  stripes(s, { x: 7.469, y: 0, w: 2.002, h: 2.001, fill: { color: GREY, transparency: 80 } });
}

// ── slide 9 · 12 step list ──────────────────────────────────────────────────
const TWELVE_STEPS = [
  ['Acknowledge the Need for Change: ', 'Recognize current limitations and the necessity for transformation.'],
  ['Identify Desired Outcomes: ', 'Detail the goals and objectives to be achieved through change.'],
  ['Understand Organizational Structure: ', "Review the company's hierarchy and decision-making processes."],
  ['Assess Organizational Culture: ', 'Evaluate how the existing culture may impact change implementation.'],
  ['Define Stakeholders: ', 'Identify all individuals and groups who will be affected by the change.'],
  ['Evaluate Available Resources: ', 'Understand the resources and skills at disposal for managing the change.'],
  ['Analyze Employee Dynamics: ', 'Consider the attitudes and expectations of the employees towards the proposed change.'],
  ['Review Past Change Experiences: ', 'Reflect on previous change initiatives and their outcomes'],
  ['Research Suitable Models: ', 'Investigate various change management models and their applications.'],
  ['Align with Strategic Vision: ', "Ensure the chosen model supports the company's long-term goals."],
  ['Plan Implementation: ', 'Develop a roadmap for implementing the selected change management model.'],
  ['Prepare for Resistance: ', 'Develop strategies to handle potential resistance to change.']
];
function slide09(pptx) {
  const s = page(pptx, 'light');
  txt(s, 'A 12-Step Process to Selecting the Optimal Change Management Model',
    { x: 0.612, y: 0.474, w: 3.732, h: 1.178, fontSize: 23 });
  txt(s, "PLACEHOLDER" +
    "organization's unique needs and characteristics. By carefully going through these 12 steps, you can ensure that the model you " +
    'choose will effectively support your business through its transformation, align with your strategic vision, and set the stage for successful change.',
    { x: 5.273, y: 0.581, w: 4.146, h: 1.027, fontSize: 10, color: GREY });
  TWELVE_STEPS.forEach(function (st, i) {
    const col = i < 6 ? 0 : 1, row = i % 6;
    const x = col ? 5.271 : 0.641, y = 1.999 + row * 0.5158;
    rrect(s, { x: x, y: y, w: 0.281, h: 0.281, adj: 24822, fill: { color: LIME } });
    txt(s, '/' + ('0' + (i + 1)).slice(-2), { x: x, y: y + 0.033, w: 0.265, h: 0.194, fontSize: 9, align: 'center' });
    txt(s, [{ text: st[0] }, { text: st[1], options: { color: GREY } }],
      { x: x + 0.435, y: y - 0.047, w: 3.626, h: 0.37, fontSize: 10 });
  });
}

// ── slides 10 & 11 · condensed approach ─────────────────────────────────────
const CONDENSED_TITLE = 'PLACEHOLDER';
const CONDENSED = [
  ['Recognize Organizational Requirements', 'Understand the unique needs, challenges, and goals of your organization. ' +
    'This insight is crucial in determining the type of change model that will best suit your company.'],
  ['Consider Cultural Compatibility', 'PLACEHOLDER' +
    "organization's culture. The model must resonate with your company's values, beliefs, and practices to ensure effective implementation and acceptance."]
];
function slide10(pptx) {
  const s = page(pptx, 'dark');
  photo(s, { x: 0.667, y: 0.639, w: 5.458, h: 1.914, adj: 13039 });
  txt(s, CONDENSED_TITLE, { x: 6.537, y: 0.572, w: 2.734, h: 1.935, fontSize: 23, color: LIGHT });
  CONDENSED.forEach(function (c, i) {
    const x = 0.835 + i * 2.771;
    txt(s, c[0], { x: x, y: 2.88, w: 2.337, h: 0.547, fontSize: 15, color: LIGHT });
    ln(s, { x: x + 0.023, y: 3.647, w: 2.364, h: 0, line: { color: GREY, width: 1.5, transparency: 70 } });
    txt(s, c[1], { x: x + 0.015, y: 3.819, w: 2.382, h: 1.027, fontSize: 10, color: GREY });
  });
  arrow(s, { x: 6.626, y: 2.915, w: 0.349, h: 0.478, rotate: -90, fill: { color: LIME } });
  stripes(s, { x: 9.537, y: 0, w: 0.463, h: 5.625, fill: { color: GREY, transparency: 70 } });
}
function slide11(pptx) {
  const s = page(pptx, 'light');
  photo(s, { x: 3.861, y: 0.639, w: 2.819, h: 4.361, adj: 8785 });
  txt(s, CONDENSED_TITLE, { x: 0.662, y: 0.572, w: 2.734, h: 1.935, fontSize: 23 });
  txt(s, 'The process of choosing a suitable Change Management Model, while streamlined in this approach, remains a strategic task. ' +
    "It primarily involves a deep understanding of your organization's unique requirements and the cultural compatibility of the model. " +
    "This focused method ensures the selected model supports your business's transformation and leads to successful change initiatives.",
    { x: 0.631, y: 2.736, w: 2.734, h: 1.519, fontSize: 10, color: GREY });
  CONDENSED.forEach(function (c, i) {
    const y = i ? 2.88 : 0.634;
    rrect(s, { x: 7.036, y: y + 0.079, w: 0.137, h: 0.137, adj: 25312, fill: { color: LIME2 } });
    txt(s, c[0], { x: 7.291, y: y, w: 1.854, h: 0.547, fontSize: 15 });
    txt(s, c[1], { x: 7.305, y: y + 0.938, w: 1.861, h: 1.355, fontSize: 10, color: GREY });
  });
  rrect(s, { x: 0.644, y: 4.461, w: 1.502, h: 0.408, adj: 27312, fill: { color: LIME } });
  txt(s, 'More Information', { x: 0.716, y: 4.566, w: 1.348, h: 0.198, fontSize: 9, align: 'center' });
  stripes(s, { x: 9.537, y: 0, w: 0.463, h: 5.625, fill: { color: GREY, transparency: 80 } });
}

// ── slides 12-14 · gallery ──────────────────────────────────────────────────
const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore ' +
  'magna aliqua. Pretium quam vulputate dignissim suspendisse in est ante in nibh. Tempor commodo ullamcorper a lacus vestibulum sed arcu.';
// lime rounded badge holding a black "+" glyph
function plusBadge(s, o) {
  rrect(s, { x: o.x, y: o.y, w: 0.418, h: 0.408, adj: 27312, fill: { color: LIME } });
  poly(s, { x: o.x + 0.138, y: o.y + 0.128, w: 0.141, h: 0.141, fill: { color: BLACK } },
    [[[0.42, 0], [0.58, 0], [0.58, 0.42], [1, 0.42], [1, 0.58], [0.58, 0.58], [0.58, 1], [0.42, 1], [0.42, 0.58], [0, 0.58], [0, 0.42], [0.42, 0.42]]]);
}
function slide12(pptx) {
  const s = page(pptx, 'light');
  photo(s, { x: 0.659, y: 0.74, w: 3.252, h: 2.593, adj: 9576 });
  photo(s, { x: 4.213, y: 0.74, w: 1.475, h: 3.912, adj: 16667 });
  photo(s, { x: 0.659, y: 3.619, w: 1.475, h: 1.033, adj: 25205 });
  photo(s, { x: 2.427, y: 3.619, w: 1.475, h: 1.033, adj: 25205 });
  txt(s, LOREM_LONG, { x: 6.137, y: 0.79, w: 3.168, h: 0.863, fontSize: 10, color: GREY });
  ln(s, { x: 6.159, y: 1.928, w: 3.165, h: 0, line: { color: BLACK, width: 1.5 } });
  txt(s, 'Gallery', { x: 6.102, y: 3.875, w: 2.589, h: 0.989, fontSize: 56 });
  txt(s, '04', { x: 8.911, y: 3.875, w: 1.212, h: 0.989, fontSize: 56, color: LIME, align: 'right' });
}
function slide13(pptx) {
  const s = page(pptx, 'light');
  photo(s, { x: 0.666, y: 0.735, w: 4.785, h: 2.916, adj: 8598 });
  photo(s, { x: 5.758, y: 0.735, w: 2.926, h: 2.916, adj: 8345 });
  photo(s, { x: 8.994, y: 0.735, w: 1.617, h: 2.916, adj: 16667 });
  plusBadge(s, { x: 0.829, y: 0.905 });
  txt(s, '04', { x: 0.578, y: 3.818, w: 1.212, h: 0.989, fontSize: 56, color: LIME });
  txt(s, 'Gallery', { x: 1.685, y: 3.818, w: 2.589, h: 0.989, fontSize: 56 });
  arrow(s, { x: 8.189, y: 4.076, w: 0.406, h: 0.557, rotate: -90, fill: { color: BLACK } });
  stripes(s, { x: 8.965, y: 3.938, w: 1.059, h: 0.762, fill: { color: GREY, transparency: 80 } });
}
function slide14(pptx) {
  const s = page(pptx, 'light');
  for (let i = 0; i < 6; i++) photo(s, { x: -0.582 + i * 1.91, y: 0.731, w: 1.617, h: 2.916, adj: 16667 });
  plusBadge(s, { x: 1.501, y: 0.905 });
  plusBadge(s, { x: 7.236, y: 0.905 });
  stripes(s, { x: -0.097, y: 3.938, w: 4.95, h: 0.762, fill: { color: GREY, transparency: 80 } });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ' +
    'Pretium quam vulputate dignissim suspendisse in est ante in nibh.',
    { x: 5.39, y: 3.925, w: 3.302, h: 0.698, fontSize: 10, color: GREY, align: 'right' });
}

// ── slide 15 · project bar chart ────────────────────────────────────────────
const CHART_DATA = [
  { month: 'April', dark: 27, lime: 74 },
  { month: 'May', dark: 15, lime: 47 },
  { month: 'June', dark: 43, lime: 48 },
  { month: 'July', dark: 12, lime: 24 },
  { month: 'August', dark: 27, lime: 93 }
];
function slide15(pptx) {
  const s = page(pptx, 'light');
  [0.655, 1.647, 2.638].forEach(function (x) { photo(s, { x: x, y: 4.073, w: 0.846, h: 0.844, adj: 16667 }); });
  photo(s, { x: 0.654, y: 0.624, w: 2.83, h: 3.285, adj: 9401 });
  plusBadge(s, { x: 0.808, y: 0.782 });
  txt(s, 'Project', { x: 4.055, y: 0.519, w: 2.723, h: 0.989, fontSize: 56 });
  txt(s, '05', { x: 8.984, y: 0.534, w: 1.104, h: 0.989, fontSize: 56, color: LIME, align: 'right' });
  txt(s, LOREM_LONG, { x: 4.112, y: 1.495, w: 5.134, h: 0.534, fontSize: 10, color: GREY });
  ln(s, { x: 4.128, y: 2.371, w: 5.197, h: 0, line: { color: BLACK, width: 1.5 } });
  // plot area: y axis 0..100 maps to 4.65 .. 2.778
  const base = 4.65, top = 2.778, scale = (base - top) / 100;
  rect(s, { x: 4.586, y: 2.816, w: 1.827, h: 1.833, fill: { color: LIME2, transparency: 78 } });
  [0, 1, 2, 3].forEach(function (i) {
    ln(s, { x: 4.444, y: 2.818 + i * 0.4583, w: 4.86, h: 0, line: { color: GREY, width: 0.75, transparency: 50, dashType: 'dash' } });
  });
  ['100', '75', '50', '25', '0'].forEach(function (v, i) {
    txt(s, v, { x: 3.991, y: 2.718 + i * 0.4385, w: 0.375, h: 0.21, fontSize: 10, color: GREY, align: 'right', valign: 'middle' });
  });
  CHART_DATA.forEach(function (d, i) {
    const x = 4.81 + i * 0.9693;
    rect(s, { x: x, y: base - d.dark * scale, w: 0.119, h: d.dark * scale, fill: { color: BLACK } });
    rect(s, { x: x + 0.161, y: base - d.lime * scale, w: 0.119, h: d.lime * scale, fill: { color: LIME } });
    txt(s, d.month, { x: x - 0.216, y: 4.681, w: 0.665, h: 0.21, fontSize: 10, color: GREY, align: 'center', valign: 'middle' });
  });
  // axis lines + the marker dropping onto June
  ln(s, { x: 4.44, y: 2.778, w: 0, h: 1.872, line: { color: GREY, width: 1.5 } });
  ln(s, { x: 4.44, y: 4.65, w: 4.859, h: 0, line: { color: GREY, width: 1.5 } });
  ln(s, { x: 6.414, y: 2.663, w: 0, h: 2.169, line: { color: BLACK, width: 1.5, dashType: 'dash' } });
  ell(s, { x: 6.285, y: 2.475, w: 0.259, h: 0.259, fill: { color: BLACK } });
  tick(s, { x: 6.334, y: 2.552, w: 0.153, h: 0.107, fill: { color: LIME } });
}

// ── slide 16 · testimonial ──────────────────────────────────────────────────
function slide16(pptx) {
  const s = page(pptx, 'dark');
  photo(s, { x: 0, y: 0, w: 4.982, h: 5.625, adj: 0 });
  stripes(s, { x: -0.005, y: 0, w: 0.474, h: 5.625, fill: { color: GREY, transparency: 70 } });
  rrect(s, { x: 3.766, y: 0.683, w: 5.461, h: 4.091, adj: 6365, fill: { color: GREY, transparency: 70 } });
  txt(s, 'PLACEHOLDER',
    { x: 4.204, y: 1.072, w: 4.045, h: 1.178, fontSize: 23, color: LIGHT });
  txt(s, 'We engaged the services of this consulting agency to develop a Change Management Model for our organization, and the results ' +
    "were beyond our expectations. Their team's expertise in analyzing our needs, understanding our culture, and crafting a tailored " +
    'approach has been instrumental in our successful business transformation. The change process was seamless, and our team was ' +
    'highly engaged throughout. I highly recommend their services for any organization looking to navigate change effectively.',
    { x: 4.216, y: 2.415, w: 4.708, h: 1.191, fontSize: 10, color: GREY });
  rrect(s, { x: 8.606, y: 0.874, w: 0.418, h: 0.408, adj: 27312, fill: { color: LIME } });
  quote(s, { x: 8.727, y: 0.999, w: 0.185, h: 0.157, color: BLACK });
  txt(s, [{ text: 'John Smith', options: { color: LIME, breakLine: true } },
    { text: 'CEO, Tech Innovations Ltd.', options: { color: GREY, breakLine: true } },
    { text: '     456 Silicon Valley Blvd, San Jose, California, USA', options: { color: GREY } }],
    { x: 4.216, y: 3.789, w: 4.417, h: 0.534, fontSize: 10 });
  shieldTick(s, { x: 4.234, y: 4.173, w: 0.108, h: 0.132 });
  // hand-written signature
  const sig = [];
  for (let i = 0; i <= 40; i++) {
    const t = i / 40;
    sig.push({ x: 7.131 + 1.691 * t, y: 3.587 + 0.897 * (0.55 + 0.42 * Math.sin(t * 12) * Math.exp(-2.2 * t) - 0.30 * Math.exp(-6 * t)) });
  }
  s.addShape('custGeom', { x: 7.131, y: 3.587, w: 1.691, h: 0.897, points: sig, line: { color: LIME, width: 1.25 } });
}

// ── slide 17 · three testimonials ───────────────────────────────────────────
const TESTIMONIALS = [
  ['A Game-Changer in Change Management', 'Our experience with this consulting agency was transformative. Their ability to tailor a ' +
    'Change Management Model to our needs was impressive and effective.', 'Sarah Mitchell', 'Chief Operations Officer', 'Pacific Technologies Inc.'],
  ['Expert Guidance Through Change', 'This team provided us with expert guidance, facilitating a smooth transition for our organization. ' +
    'Their Change Management Model was the roadmap we needed.', 'Mark Daniels', 'Director of Human Resources', 'State Manufacturing'],
  ['Tailored Approach to Effective Change', 'The consulting agency took the time to understand our organization, developing a Change ' +
    'Management Model that was perfectly aligned with our needs.', 'Laura Gomez', 'Executive Director', 'California Health Systems']
];
function slide17(pptx) {
  const s = page(pptx, 'grey');
  txt(s, 'Testimonial', { x: 0.612, y: 0.485, w: 3.661, h: 0.673, fontSize: 38 });
  txt(s, '06', { x: -0.107, y: 0.485, w: 0.779, h: 0.673, fontSize: 38, color: LIME });
  TESTIMONIALS.forEach(function (t, i) {
    const x = 0.662 + i * 2.995;
    rrect(s, { x: x, y: 1.533, w: 2.679, h: 3.052, adj: 9722, fill: { color: BLACK } });
    txt(s, t[0], { x: x + 0.293, y: 1.846, w: 2.166, h: 0.421, fontSize: 11, color: LIGHT });
    quote(s, { x: x + 2.274, y: 1.729, w: 0.185, h: 0.157, color: GREY });
    txt(s, t[1], { x: x + 0.293, y: 2.315, w: 2.166, h: 0.863, fontSize: 10, color: GREY });
    for (let k = 0; k < 5; k++) s.addShape('star5', { x: x + 0.901 + k * 0.167, y: 3.664, w: 0.11, h: 0.105, fill: { color: SOFT } });
    photo(s, { x: x + 0.222, y: 3.85, w: 0.503, h: 0.492, adj: 28037, color: GREY, label: LIGHT });
    txt(s, [{ text: t[2], options: { color: LIME, breakLine: true } },
      { text: t[3], options: { color: GREY, breakLine: true } },
      { text: '     ' + t[4], options: { color: GREY } }],
      { x: x + 0.871, y: 3.84, w: 1.669, h: 0.534, fontSize: 10 });
    shieldTick(s, { x: x + 0.89, y: 4.22, w: 0.108, h: 0.132 });
  });
}

// ── slide 18 · models comparison ────────────────────────────────────────────
// each row = grey track + label chip + a run of homePlate segments
const MODEL_ROWS = [
  { label: 'Lewin', cells: [[1.652, 2.510, BLACK, 'Unfreezing'], [4.211, 2.510, GREY, 'Change and Transition'], [6.777, 2.510, LIME, 'Refreezing']] },
  { label: 'Prosci\nADKAR', cells: [[1.652, 1.240, BLACK, 'Awareness'], [2.934, 1.229, BLACK, 'Desire'], [4.211, 1.234, GREY, 'Knowledge'], [5.487, 1.234, GREY, 'Ability'], [6.777, 2.510, LIME, 'Reinforcement']] },
  { label: 'Stages', cells: [[1.652, 1.240, BLACK, 'Precontemplation'], [2.934, 1.229, BLACK, 'Contemplation'], [4.211, 1.234, GREY, 'Preparation'], [5.487, 1.234, GREY, 'Action'], [6.777, 2.510, LIME, 'Maintenance']] },
  { label: 'Stages', cells: [[1.652, 0.911, BLACK, 'Urgency'], [2.612, 0.911, BLACK, 'Coalition'], [3.573, 0.911, BLACK, 'Create vision'], [4.534, 0.911, GREY, 'Comm. vision'], [5.495, 0.911, GREY, 'Obstacles'], [6.456, 0.911, LIME, 'Short-term wins'], [7.416, 0.911, LIME, 'Build'], [8.377, 0.911, LIME, 'Anchor']] }
];
function slide18(pptx) {
  const s = page(pptx, 'light');
  txt(s, 'Different Models and Theories of Change Management', { x: 0.631, y: 0.515, w: 7.316, h: 0.421, fontSize: 23 });
  txt(s, [{ text: 'Change management (sometimes abbreviated as CM) is a collective term for all approaches to prepare, support, and help ' +
      'individuals, teams, and organizations in making organizational change. It includes methods that redirect or redefine the use of resources, ' +
      'business process, budget allocations, or other modes of operation that significantly change a company or organization.', options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: 'Organizational change management (OCM) considers the full organization and what needs to change, while change management may be ' +
      'used solely to refer to how people and teams are affected by such organizational transition. It deals with many different disciplines, from ' +
      'behavioral and social sciences to information technology and business solutions.' }],
    { x: 0.653, y: 1.083, w: 8.532, h: 1.191, fontSize: 10, color: GREY });
  MODEL_ROWS.forEach(function (row, r) {
    const y = 2.604 + r * 0.5707;
    rrect(s, { x: 0.672, y: y, w: 8.642, h: 0.521, adj: 10605, fill: { color: GREY, transparency: 95 } });
    rrect(s, { x: 0.724, y: y + 0.053, w: 0.878, h: 0.417, adj: 6716, fill: { color: GREY, transparency: 80 } });
    txt(s, row.label, { x: 0.724, y: y + 0.09, w: 0.878, h: 0.34, fontSize: 8, align: 'center', valign: 'middle' });
    row.cells.forEach(function (c) {
      tag(s, 'homePlate', { x: c[0], y: y + 0.053, w: c[1], h: 0.417, point: 0.13, fill: { color: c[2] } });
      txt(s, c[3], { x: c[0], y: y + 0.09, w: c[1] - 0.13, h: 0.34, fontSize: 8, align: 'center', valign: 'middle', color: c[2] === LIME ? GREY : LIGHT });
    });
  });
}

// ── slide 19 · transformation gear ──────────────────────────────────────────
const GEAR_LABELS = [
  ['Organisation Review', 3, 2.428, 1.418, 1.617, 'r'], ['Operating Model', 2, 1.975, 1.738, 1.617, 'r'],
  ['Governance', 1, 1.650, 2.203, 1.617, 'r'], ['Transition Planning', 3, 1.435, 3.625, 1.835, 'r'],
  ['Organisational Readiness', 2, 1.758, 4.085, 1.835, 'r'], ['Transition Management Support', 1, 1.886, 4.416, 2.181, 'r'],
  ['Process Identification', 3, 5.939, 1.421, 1.617, 'l'], ['Process Optimization', 2, 6.410, 1.740, 1.617, 'l'],
  ['Process Implementation', 1, 6.717, 2.205, 1.617, 'l'], ['Change Management Planning', 3, 6.717, 3.627, 1.835, 'l'],
  ['Change Delivery', 2, 6.401, 4.087, 1.835, 'l'], ['Change Education Programme', 1, 5.925, 4.419, 1.835, 'l']
];
const GEAR_BADGES = [[3.390, 2.201], [6.362, 2.201], [5.540, 4.399], [4.229, 4.399], [3.729, 4.066], [6.019, 4.066],
  [3.728, 1.736], [6.018, 1.736], [3.405, 3.610], [6.357, 3.610], [4.220, 1.395], [5.547, 1.395]];
const GEAR_BADGE_NUM = ['1', '1', '1', '1', '2', '2', '2', '2', '3', '3', '3', '3'];
function slide19(pptx) {
  const s = page(pptx, 'lime');
  stripes(s, { x: -0.011, y: 1.7, w: 10.016, h: 2.614, fill: { color: LIGHT, transparency: 70 } });
  txt(s, 'Business Transformation and Change Management', { x: 1.347, y: 0.567, w: 7.316, h: 0.421, fontSize: 23, align: 'center' });
  // "gear" = pale disc; the 12 lime badges below straddle its rim and notch it
  ell(s, { x: 3.349, y: 1.371, w: 3.302, h: 3.302, fill: { color: SOFT } });
  // four black quadrants (pie wedges) inside a black ring, split by light gaps
  ell(s, { x: 3.680, y: 1.696, w: 2.650, h: 2.650, fill: { color: BLACK } });
  [0, 90, 180, 270].forEach(function (a) {
    s.addShape('pie', { x: 3.694, y: 1.710, w: 2.622, h: 2.622, angleRange: [a + 2, a + 88], fill: { color: BLACK }, line: { color: SOFT, width: 1.5 } });
  });
  ell(s, { x: 3.650, y: 1.666, w: 2.710, h: 2.710, fill: { color: SOFT, transparency: 100 }, line: { color: BLACK, width: 3 } });
  ell(s, { x: 4.470, y: 2.491, w: 1.061, h: 1.061, fill: { color: SOFT } });
  [['Organizational\nEffectiveness', 3.906, 2.264], ['Business Process Excellence', 5.160, 2.264],
    ['Transition\nManagement', 3.929, 3.507], ['Change Management', 5.156, 3.507]].forEach(function (q) {
    txt(s, q[0], { x: q[1], y: q[2], w: 0.878, h: 0.32, fontSize: 8, color: LIGHT, align: 'center' });
  });
  txt(s, 'Transformation\n& Change', { x: 4.566, y: 2.876, w: 0.878, h: 0.295, fontSize: 8, align: 'center' });
  GEAR_LABELS.forEach(function (g) {
    txt(s, g[0], { x: g[2], y: g[3], w: g[4], h: 0.214, fontSize: 10, color: GREY, align: g[5] === 'r' ? 'right' : 'left' });
  });
  GEAR_BADGES.forEach(function (b, i) {
    ell(s, { x: b[0] - 0.16, y: b[1] - 0.16, w: 0.567, h: 0.567, fill: { color: LIME } });
    txt(s, GEAR_BADGE_NUM[i], { x: b[0] - 0.009, y: b[1] + 0.025, w: 0.265, h: 0.198, fontSize: 9, align: 'center' });
  });
}

// ── slide 20 · Lewin three stages ───────────────────────────────────────────
const LEWIN = [
  { chip: 'Unfreezing', icon: '✹', head: 'Readiness to Change', items: ['Educate (everyone understands)', 'Inform (what, why, when, how)',
    'Consult (seek views and ideas, allow thinking time, use others’ ideas)', 'Plan (objectives, resources, time-scales, measures, budgets)',
    'Organise (work plans)', 'Appoint (leaders, managers, teams)'] },
  { chip: 'Changing', icon: '✥', head: 'Implementation', items: ['Praise', 'Encouragement', 'Recognition and Empathy', 'Coach', 'Train',
    'Lead', 'Manage', 'Help and Guidance', 'Regular feedback', 'Provide adequate resources'] },
  { chip: 'Making it stick', icon: '❉', head: 'Making it stick', items: ['Set performance indicators', 'Monitor and evaluate performances',
    'Establish systems to make it happen', 'Establish controls to check it is happening', 'Rewards for new behaviour',
    'Sanction (or lack of reward) for old behaviour', 'A period of relative stability (consolidate)'] }
];
function slide20(pptx) {
  const s = page(pptx, 'light');
  txt(s, 'Lewin’s Three Stage of Change Model:\nUnfreezing, Changing & Refreezing', { x: 0.631, y: 0.515, w: 7.316, h: 0.799, fontSize: 23 });
  txt(s, 'Kurt Lewin developed a change model involving three steps: unfreezing, changing and refreezing. For Lewin, the process of change ' +
    'entails creating the perception that a change is needed, then moving toward the new, desired level of behavior and, finally, solidifying ' +
    'that new behavior as the norm.', { x: 0.648, y: 1.473, w: 8.532, h: 0.37, fontSize: 10, color: GREY });
  LEWIN.forEach(function (c, i) {
    const x = 0.673 + i * 2.951;
    rrect(s, { x: x, y: 2.17, w: 2.608, h: 2.734, adj: 9984, fill: { color: BLACK } });
    rrect(s, { x: x + 0.144, y: 2.31, w: 2.319, h: 0.408, adj: 30038, fill: { color: LIME } });
    txt(s, c.icon + '  ' + c.chip, { x: x + 0.144, y: 2.31, w: 2.319, h: 0.408, fontSize: 9, align: 'center', valign: 'middle' });
    txt(s, c.head, { x: x + 0.418, y: 2.884, w: 1.771, h: 0.219, fontSize: 11, color: LIGHT, align: 'center' });
    txt(s, c.items.map(function (t) { return { text: t, options: { breakLine: true } }; }),
      { x: x + 0.219, y: 3.222, w: 2.17, h: 1.304, fontSize: 8, color: GREY, bullet: { characterCode: '002D', indent: 8 } });
  });
}

// ── slide 21 · Bridges transition model ─────────────────────────────────────
const BRIDGES = [
  { top: 'Ending', bottom: 'Reconciliation', x: 0.677 },
  { top: 'Transition', bottom: 'Reorientation', x: 3.573 },
  { top: 'New Beginnings', bottom: 'Recommitment', x: 6.469 }
];
const BRIDGES_WORDS = [['Denial', 0.546, 2.203], ['Anxiety', 1.088, 2.341], ['Shock', 1.567, 2.516], ['Fear', 1.953, 2.703],
  ['Anger', 2.307, 2.911], ['Frustration', 2.682, 3.234], ['Confusion', 3.161, 3.568], ['Stress', 3.650, 3.797],
  ['Approach - Avoidance', 4.088, 4.026, 1.814], ['Creativity', 5.416, 3.797], ['Skepticism', 6.005, 3.568],
  ['Acceptance', 6.469, 3.234], ['Impatience', 6.817, 2.911], ['Hope', 7.338, 2.513], ['Energy', 7.828, 2.341], ['Enthusiasm', 8.453, 2.263]];
function slide21(pptx) {
  const s = page(pptx, 'light');
  txt(s, 'William Bridges Transition Model', { x: 0.621, y: 0.546, w: 7.316, h: 0.421, fontSize: 23 });
  BRIDGES.forEach(function (c) {
    rrect(s, { x: c.x, y: 1.685, w: 2.843, h: 2.709, adj: 4806, fill: { color: GREY, transparency: 90 } });
    rrect(s, { x: c.x, y: 1.293, w: 2.843, h: 0.339, adj: 30762, fill: { color: LIME } });
    txt(s, c.top, { x: c.x, y: 1.364, w: 2.843, h: 0.214, fontSize: 10, color: GREY, align: 'center' });
    rrect(s, { x: c.x, y: 4.447, w: 2.843, h: 0.339, adj: 30762, fill: { color: BLACK } });
    txt(s, c.bottom, { x: c.x, y: 4.512, w: 2.843, h: 0.214, fontSize: 10, color: LIGHT, align: 'center' });
  });
  ln(s, { x: 0.695, y: 2.719, w: 8.61, h: 0, line: { color: GREY, width: 1.5, transparency: 70, dashType: 'dash' } });
  // the lime "dip" curve running across all three panels
  s.addShape('custGeom', { x: 0.677, y: 2.066, w: 8.627, h: 1.787, line: { color: LIME, width: 4 },
    points: [{ x: 0, y: 0.010 },
      { x: 1.990, y: 0.541, curve: { type: 'cubic', x1: 0.705, y1: -0.010, x2: 1.395, y2: 0.176 } },
      { x: 4.313, y: 1.787, curve: { type: 'cubic', x1: 2.751, y1: 1.008, x2: 3.394, y2: 1.782 } },
      { x: 6.661, y: 0.545, curve: { type: 'cubic', x1: 5.239, y1: 1.792, x2: 5.893, y2: 1.016 } },
      { x: 8.627, y: 0.000, curve: { type: 'cubic', x1: 7.249, y1: 0.184, x2: 7.929, y2: -0.007 } }] });
  BRIDGES_WORDS.forEach(function (w) {
    txt(s, w[0], { x: w[1], y: w[2], w: w[3] || 1.001, h: 0.214, fontSize: 10, color: GREY, align: 'center' });
  });
  // two half-circle arcs with arrow heads forming the "cycle" ring
  [[190, 350], [10, 170]].forEach(function (a) {
    s.addShape('arc', { x: 4.029, y: 1.755, w: 1.932, h: 1.928, angleRange: a,
      line: { color: BLACK, width: 1.5, endArrowType: 'triangle' } });
  });
  txt(s, 'Normal Productivity', { x: 4.494, y: 2.25, w: 1.001, h: 0.408, fontSize: 11, align: 'center' });
}

// ── slide 22 · CHAMPS2 phases ───────────────────────────────────────────────
const CHAMPS = [
  ['Phase 0', 'Transformation Initiation', BLACK, LIGHT], ['Phase 1', 'Visioning', LIME, GREY],
  ['Phase 2', 'Shaping and planning', LIME, GREY], ['Phase 3', 'Design', GREY, LIGHT],
  ['Phase 4', 'Service creation and realization', GREY, LIGHT], ['Phase 5', 'Proving and transition', GREY, LIGHT],
  ['Phase 6', 'Stabilization', GREY, LIGHT], ['Phase 7', 'Benefits realization', GREY, LIGHT]
];
const CHAMPS_NOTES = [
  [1.020, 2.460, 'The first phases take the organization through the development of the business case, from the strategic, outline, and full. ' +
    'This provides an organization with a comprehensive business case, supported by detailed outcomes, plus the costs and benefits of the new approach. '],
  [4.209, 2.460, 'In the middle phases of the CHAMPS2 program, the solution is designed, developed, tested, and improved before it is fully implemented. ' +
    'Throughout these phases, the outcomes and benefits models are continually refined and beneficial ownership is then allocated to ensure specific ' +
    'individuals are tasked with the responsibility for achieving these benefits.'],
  [7.029, 2.319, 'The last two phases of CHAMPS2, arguably the most crucial ones focus heavily on realizing these benefits. Monitoring and alerting ' +
    'are used tp raise awareness of any problrms in achieving benefits delivery, allowing the organization to undertake rapid review and reallocation ' +
    'of resources or skills as required to ensure the benefits realization gets back on track.']
];
function slide22(pptx) {
  const s = page(pptx, 'light');
  txt(s, 'CHAMPS2® Business Change Method', { x: 0.631, y: 0.608, w: 5.302, h: 0.443, fontSize: 23 });
  txt(s, "CHAMPS2® is a vision led, benefits driven business transformation method which is broad in scope and encompasses the whole business " +
    "change journey. It helps you define your organisation's strategic needs, and then provides a tailored route to ensure that the desired outcomes are achieved.",
    { x: 0.653, y: 1.177, w: 8.532, h: 0.37, fontSize: 10, color: GREY });
  CHAMPS.forEach(function (c, i) {
    const x = 0.681 + i * 1.061, faded = i >= 6;
    txt(s, c[0], { x: x + 0.265, y: 2.027, w: 0.951, h: 0.168, fontSize: 8 });
    tag(s, 'chevron', { x: x, y: 2.295, w: 1.211, h: 0.714, point: 0.19, fill: { color: c[2], transparency: faded ? 60 : 0 } });
    txt(s, c[1], { x: x + 0.261, y: 2.375, w: 0.951, h: 0.553, fontSize: 8, color: c[3], valign: 'middle' });
  });
  [3.850, 7.028].forEach(function (x) {
    ln(s, { x: x, y: 3.052, w: 0, h: 0.238, line: { color: GREY, width: 0.75, transparency: 60, dashType: 'dash' } });
  });
  // downward braces joining each phase group to its note
  [[0.942, 3.152], [4.130, 3.125], [7.318, 2.271]].forEach(function (b) {
    s.addShape('rightBrace', { x: b[0] + b[1] / 2 - 0.154, y: 3.29 - b[1] / 2 + 0.154, w: 0.308, h: b[1], rotate: 90,
      line: { color: GREY, width: 0.75, transparency: 60 } });
  });
  CHAMPS_NOTES.forEach(function (n) {
    txt(s, n[2], { x: n[0], y: 3.681, w: n[1], h: 1.153, fontSize: 8, color: GREY, align: 'center' });
  });
}

// ── slide 23 · shift & drift ────────────────────────────────────────────────
const DRIFT_NODES = [
  ['Analyze the data', 0.674, 2.026, LIME, GREY], ['Collect data', 0.674, 2.859, LIME, GREY], ['Drive change forward', 0.674, 3.693, LIME, GREY],
  ['Evaluate', 2.409, 1.337, GREY, LIGHT], ['Examine', 4.337, 1.337, GREY, LIGHT], ['Enumerate', 6.264, 1.337, GREY, LIGHT],
  ['Construct', 7.997, 2.026, LIME, GREY], ['Demonstrate', 7.997, 2.859, LIME, GREY], ['Determine obstacles', 7.997, 3.693, LIME, GREY],
  ['Preserve change focus', 2.409, 4.384, GREY, LIGHT], ['Manage obstacles', 4.337, 4.384, GREY, LIGHT], ['List obstacles', 6.264, 4.384, GREY, LIGHT],
  ['Explain', 4.337, 2.026, LIME, BLACK], ['Exhibit', 2.409, 2.858, LIME, BLACK], ['Educate', 6.264, 2.858, LIME, BLACK], ['Encourage', 4.337, 3.693, LIME, BLACK]
];
// Connectors. Straight ones give a direction; elbows give the side the vertical
// leg sits on (v), the side the horizontal leg sits on (hz) and which ends get heads.
const DRIFT_ARROWS = [
  { x: 3.771, y: 1.487, w: 0.505, h: 0.105, dir: 'right' }, { x: 5.698, y: 1.487, w: 0.501, h: 0.105, dir: 'right' },
  { x: 3.798, y: 4.544, w: 0.505, h: 0.105, dir: 'left' }, { x: 5.728, y: 4.544, w: 0.501, h: 0.105, dir: 'left' },
  { x: 8.612, y: 2.466, w: 0.105, h: 0.330, dir: 'down' }, { x: 8.612, y: 3.303, w: 0.105, h: 0.334, dir: 'down' },
  { x: 1.278, y: 2.505, w: 0.105, h: 0.319, dir: 'up' }, { x: 1.278, y: 3.334, w: 0.105, h: 0.330, dir: 'up' },
  { x: 1.302, y: 1.487, w: 1.050, h: 0.510, v: 'l', hz: 't', heads: ['right'] },
  { x: 7.617, y: 1.513, w: 1.094, h: 0.449, v: 'r', hz: 't', heads: ['down'] },
  { x: 7.645, y: 4.135, w: 1.043, h: 0.514, v: 'r', hz: 'b', heads: ['left'] },
  { x: 1.278, y: 4.168, w: 1.099, h: 0.453, v: 'l', hz: 'b', heads: ['up'] },
  { x: 3.044, y: 2.180, w: 1.225, h: 0.651, v: 'l', hz: 't', heads: ['right', 'down'] },
  { x: 5.728, y: 2.207, w: 1.246, h: 0.584, v: 'r', hz: 't', heads: ['left', 'down'] },
  { x: 3.020, y: 3.329, w: 1.285, h: 0.599, v: 'l', hz: 'b', heads: ['right', 'up'] },
  { x: 5.727, y: 3.305, w: 1.221, h: 0.652, v: 'r', hz: 'b', heads: ['left', 'up'] },
  { x: 2.100, y: 2.210, w: 0.911, h: 0.626, v: 'r', hz: 't', heads: ['down'] },
  { x: 3.024, y: 1.480, w: 1.255, h: 0.666, v: 'l', hz: 'b', heads: ['right', 'up'] },
  { x: 5.564, y: 1.855, w: 1.390, h: 0.317, v: 'r', hz: 'b', heads: ['left'] },
  { x: 7.020, y: 2.180, w: 0.861, h: 0.635, v: 'l', hz: 't', heads: ['left', 'down'] },
  { x: 6.996, y: 3.329, w: 0.885, h: 0.599, v: 'l', hz: 'b', heads: ['left', 'up'] },
  { x: 2.127, y: 3.320, w: 0.861, h: 0.635, v: 'r', hz: 't', heads: ['right', 'down'] },
  { x: 3.526, y: 3.900, w: 1.239, h: 0.270, v: 'l', hz: 'b', heads: ['right'] },
  { x: 5.231, y: 3.900, w: 1.251, h: 0.237, v: 'r', hz: 'b', heads: ['left'] }
];
const T = 0.062; // connector stroke thickness
const HEAD_ROT = { right: 0, left: 180, up: 270, down: 90 };
function head(s, dir, cx, cy) {
  const horiz = dir === 'right' || dir === 'left';
  s.addShape('rightArrow', { x: cx - (horiz ? 0.105 : T / 2), y: cy - (horiz ? T / 2 : 0.105),
    w: horiz ? 0.21 : T, h: horiz ? T : 0.21, rotate: HEAD_ROT[dir], fill: { color: SOFT } });
}
function flowArrow(s, a) {
  if (a.dir) {
    const horiz = a.dir === 'right' || a.dir === 'left';
    s.addShape('rightArrow', { x: a.x, y: a.y, w: horiz ? a.w : a.h, h: horiz ? a.h : a.w,
      rotate: HEAD_ROT[a.dir], fill: { color: SOFT } });
    return;
  }
  const vx = a.v === 'l' ? a.x : a.x + a.w - T;
  const hy = a.hz === 't' ? a.y : a.y + a.h - T;
  rect(s, { x: a.x, y: hy, w: a.w, h: T, fill: { color: SOFT } });
  rect(s, { x: vx, y: a.y, w: T, h: a.h, fill: { color: SOFT } });
  a.heads.forEach(function (d) {
    if (d === 'right') head(s, d, a.x + a.w, hy + T / 2);
    else if (d === 'left') head(s, d, a.x, hy + T / 2);
    else if (d === 'up') head(s, d, vx + T / 2, a.y);
    else head(s, d, vx + T / 2, a.y + a.h);
  });
}
function slide23(pptx) {
  const s = page(pptx, 'dark');
  txt(s, 'Organizational Change Management Models: Shift and drift changes',
    { x: 0.543, y: 0.567, w: 8.908, h: 0.421, fontSize: 23, color: LIGHT, align: 'center' });
  ell(s, { x: 3.497, y: 1.585, w: 3.006, h: 3.006, fill: { color: GREY, transparency: 70 } });
  DRIFT_ARROWS.forEach(function (a) { flowArrow(s, a); });
  rrect(s, { x: 2.055, y: 1.807, w: 5.886, h: 2.521, adj: 5894, fill: { color: BLACK, transparency: 100 },
    line: { color: LIGHT, width: 1.5, dashType: 'dash' } });
  txt(s, 'Change', { x: 4.485, y: 2.883, w: 1.030, h: 0.370, fontSize: 19, color: LIGHT, align: 'center' });
  DRIFT_NODES.forEach(function (n) {
    rrect(s, { x: n[1], y: n[2], w: 1.327, h: 0.416, adj: 25027, fill: { color: n[3] } });
    txt(s, n[0], { x: n[1] + 0.144, y: n[2] + 0.125, w: 1.038, h: 0.161, fontSize: 7, color: n[4], align: 'center' });
  });
}

// ── slide 24 · Burke-Litwin ─────────────────────────────────────────────────
const BURKE_NODES = [
  ['External Environment', 2.625, 0.786, LIME, GREY, 0.168],
  ['Mission & Strategy', 1.065, 1.492, 'DARK40', LIGHT, 0.168],
  ['Leadership', 2.625, 1.492, 'DARK40', LIGHT, 0.168],
  ['Organizational Culture', 4.190, 1.492, 'DARK40', LIGHT, 0.168],
  ['Structure', 1.065, 2.194, LIME, GREY, 0.168],
  ['Management Practices', 2.625, 2.194, LIME, GREY, 0.168],
  ['System & Policies', 4.190, 2.194, LIME, GREY, 0.168],
  ['Work Until Climate', 2.625, 2.905, BLACK, LIGHT, 0.168],
  ['Task & Individual Skills', 1.065, 3.612, LIME, GREY, 0.168],
  ['Motivation', 2.625, 3.612, LIME, GREY, 0.168],
  ['Individual Needs & Values', 4.190, 3.612, LIME, GREY, 0.295],
  ['Individual & Organizational Performance', 2.625, 4.318, 'DARK40', LIGHT, 0.421]
];
// vertical / horizontal hairlines wiring the boxes together
const BURKE_LINES = [
  [3.314, 1.299, 0, 0.192], [3.314, 2.014, 0, 0.192], [3.314, 2.672, 0, 0.233], [3.314, 3.429, 0, 0.183], [3.314, 3.998, 0, 0.324],
  [1.753, 2.014, 0, 0.186], [1.753, 2.714, 0, 0.897], [4.879, 2.020, 0, 0.172], [4.879, 2.714, 0, 1.010],
  [2.447, 1.753, 0.165, 0], [4.014, 1.753, 0.165, 0], [2.442, 2.452, 0.181, 0], [3.957, 2.454, 0.274, 0],
  [2.442, 3.872, 0.184, 0], [3.866, 3.872, 0.460, 0]
];
function slide24(pptx) {
  const s = page(pptx, 'grey');
  // nested bracket frames wrapping the factor groups
  s.addShape('bracketPair', { x: 0.679, y: 0.918, w: 5.256, h: 3.675, rectRadius: 0.12, line: { color: SOFT, width: 0.75 } });
  s.addShape('bracketPair', { x: 0.857, y: 1.763, w: 4.908, h: 2.113, rectRadius: 0.12, line: { color: SOFT, width: 0.75 } });
  s.addShape('bracePair', { x: 1.739, y: 1.050, w: 3.150, h: 0.434, rectRadius: 0.1, rotate: 180, line: { color: SOFT, width: 0.75 } });
  BURKE_LINES.forEach(function (l) {
    ln(s, { x: l[0], y: l[1], w: l[2], h: l[3], line: { color: LIGHT, width: 0.75 } });
  });
  // crossing diagonals between leadership and its neighbours
  [[2.431, 2.007], [3.996, 2.003]].forEach(function (c) {
    ln(s, { x: c[0], y: c[1], w: 0.2, h: 0.195, line: { color: SOFT, width: 0.75 } });
    ln(s, { x: c[0], y: c[1], w: 0.2, h: 0.195, flipV: true, line: { color: SOFT, width: 0.75 } });
  });
  BURKE_NODES.forEach(function (n) {
    const dark40 = n[3] === 'DARK40';
    rrect(s, { x: n[1], y: n[2], w: 1.377, h: 0.521, adj: 20000, fill: { color: dark40 ? BLACK : n[3], transparency: dark40 ? 60 : 0 } });
    txt(s, n[0], { x: n[1] + 0.136, y: n[2] + 0.26 - n[5] / 2, w: 1.105, h: n[5], fontSize: 8, color: n[4], align: 'center', valign: 'middle' });
  });
  txt(s, 'Burke-Litwin\nChange Model', { x: 6.589, y: 1.324, w: 2.735, h: 0.799, fontSize: 23, valign: 'bottom' });
  txt(s, [{ text: 'The Burke-Litwin change model is an organizational change model. It says that there are 12 key factors that organizations ' +
      'must consider when assessing change. The model groups these factors into different levels.', options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: 'At the most macro level there are external factors. At the next level down are strategic factors, which include organizational ' +
      'culture. These are followed by operating factors and individual factors, before we reach the most micro factor grouping which is output.' }],
    { x: 6.558, y: 2.285, w: 2.735, h: 2.010, fontSize: 10, color: LIGHT });
}

// ── slide 25 · viral change roadmap ─────────────────────────────────────────
const VIRAL = [
  { phase: 'Discovery', icon: '✦', chip: LIME, chipText: BLACK, out: 'Change strategy blueprint', outColor: GREY, outFill: LIME,
    items: ['Visioning', 'Uncovering behaviors', 'Mapping influence', 'Model of social infection', 'Project leads'] },
  { phase: 'Development', icon: '⏣', chip: BLACK, chipText: LIGHT, out: 'Project management pre-launch of champions work ', outColor: LIGHT, outFill: BLACK,
    items: ['Stakeholder management', 'Champions pool', 'Top-down, non viral communication framework', 'Project management: pre-champions first conference'] },
  { phase: 'Engagement', icon: '⬢', chip: BLACK, chipText: LIGHT, out: 'Champions engaged management aligned', outColor: LIGHT, outFill: BLACK,
    items: ['Champions first conference', 'Backstage leadership '] },
  { phase: 'Diffusion', icon: '◈', chip: LIME, chipText: BLACK, out: 'New norms emerging successes trached stories broadcasted', outColor: GREY, outFill: LIME,
    items: ['Champions in actions', 'Non-viral communication campaign', 'Champions community support', 'Tracking progress', 'Review point', 'Management support'] },
  { phase: 'Sustain', icon: '◎', chip: GREY, chipText: LIGHT, out: 'Consolidated outcome redirections', outColor: LIGHT, outFill: GREY,
    items: ['Adjustments', 'Re-focus', 'Tracking progress', 'Review points: all levels', 'Extended learning', 'Viral change inside'] }
];
function slide25(pptx) {
  const s = page(pptx, 'light');
  txt(s, 'Viral Change Roadmap', { x: 0.621, y: 0.379, w: 5.897, h: 0.673, fontSize: 38 });
  // outcomes row of chevrons
  VIRAL.forEach(function (v, i) {
    const x = 1.015 + i * 1.4318;
    tag(s, 'chevron', { x: x, y: 1.376, w: 1.618, h: 0.954, point: 0.26, fill: { color: v.outFill } });
    txt(s, v.out, { x: x + 0.352, y: 1.639, w: 1.105, h: 0.421, fontSize: 8, color: v.outColor, valign: 'middle' });
  });
  // big right-pointing band behind the phase cards
  tag(s, 'homePlate', { x: 0.905, y: 2.510, w: 8.579, h: 2.389, point: 0.75, fill: { color: GREY, transparency: 90 } });
  VIRAL.forEach(function (v, i) {
    const x = 1.005 + i * 1.4835;
    rrect(s, { x: x, y: 2.787, w: 1.418, h: 0.642, adj: 4276, fill: { color: v.chip } });
    txt(s, v.icon, { x: x, y: 2.85, w: 1.418, h: 0.25, fontSize: 11, color: v.chipText, align: 'center', valign: 'middle' });
    txt(s, v.phase, { x: x + 0.157, y: 3.149, w: 1.105, h: 0.168, fontSize: 8, color: v.chipText, align: 'center', valign: 'middle' });
    txt(s, v.items.map(function (t) { return { text: t, options: { breakLine: true } }; }),
      { x: x + 0.049, y: 3.511, w: 1.392, h: 1.052, fontSize: 8, color: GREY, bullet: { characterCode: '002D', indent: 8 } });
  });
  // left-hand axis brackets labelling the two bands
  [[1.382, 0.934, 'Outcomes', 1.765], [2.692, 2.003, 'Phases', 3.636]].forEach(function (b) {
    ln(s, { x: 0.671, y: b[0], w: 0.249, h: 0, line: { color: GREY, width: 0.75, transparency: 50 } });
    ln(s, { x: 0.671, y: b[0] + b[1], w: 0.249, h: 0, line: { color: GREY, width: 0.75, transparency: 50 } });
    ln(s, { x: 0.821, y: b[0] + 0.011, w: 0, h: b[1], line: { color: GREY, width: 0.75, transparency: 50, dashType: 'dash',
      beginArrowType: 'triangle', endArrowType: 'triangle' } });
    txt(s, b[2], { x: b[3] === 1.765 ? 0.212 : 0.127, y: b[3], w: b[3] === 1.765 ? 0.934 : 1.105, h: 0.168, fontSize: 8, color: GREY,
      align: 'center', valign: 'middle', rotate: -90 });
  });
}

// ── slide 26 · Kotter 8 steps ───────────────────────────────────────────────
// [number, label, bar x, bar top, label y, colour]
const KOTTER = [
  ['01', 'Establish urgency', 1.616, 3.778, 3.400, BLACK], ['02', 'Form Leadership Coalition', 2.408, 3.619, 3.201, BLACK],
  ['03', 'Develop\nVision', 3.200, 3.461, 3.106, BLACK], ['04', 'Communicate Vision for Change', 4.498, 3.203, 2.774, LIME],
  ['05', 'Empower Employees to Participate in Change', 5.293, 3.046, 2.365, LIME], ['06', 'Create Quick Wins', 6.082, 2.890, 2.494, LIME],
  ['07', 'Stimulate Change', 7.380, 2.630, 2.254, GREY], ['08', 'Sustain\nChange ', 8.172, 2.468, 2.099, GREY]
];
const KOTTER_GROUPS = [
  [1.191, 2.381, 2.678, 2.299, 3398, 'Create a climate for change', 1.638, 4.820, 1.219, 2.624],
  [4.075, 1.854, 2.678, 2.826, 2917, 'Involve the entire company', 4.521, 4.822, 4.102, 2.624],
  [6.958, 1.244, 1.889, 3.436, 4135, 'Implement and sustain change culture', 6.950, 4.822, 6.984, 1.847]
];
function slide26(pptx) {
  const s = page(pptx, 'light');
  txt(s, "Kotter's 8-Step Change Model", { x: 0.631, y: 0.608, w: 5.302, h: 0.421, fontSize: 23 });
  txt(s, 'The 8-Step Process for Leading Change was cultivated from over four decades of Dr. Kotter’s observations of countless leaders and ' +
    'organizations as they were trying to transform or execute their strategies. He identified and extracted the success factors and combined ' +
    'them into a methodology — 8-Step Process for Leading Change.',
    { x: 0.653, y: 1.177, w: 6.339, h: 0.534, fontSize: 10, color: GREY });
  KOTTER_GROUPS.forEach(function (g) {
    rrect(s, { x: g[0], y: g[1], w: g[2], h: g[3], adj: g[4], fill: { color: GREY, transparency: 92 } });
    txt(s, g[5], { x: g[6], y: g[7], w: 1.902, h: 0.168, fontSize: 8, align: 'center' });
    ln(s, { x: g[8], y: 4.312, w: g[9], h: 0, line: { color: GREY, width: 0.75, transparency: 50, beginArrowType: 'triangle', endArrowType: 'triangle' } });
    [g[8] - 0.025, g[8] + g[9] + 0.026].forEach(function (x) {
      ln(s, { x: x, y: 4.286, w: 0, h: 0.151, line: { color: GREY, width: 0.75, transparency: 50, dashType: 'dash' } });
    });
    s.addShape('rightBrace', { x: g[8] + g[9] / 2 - 0.154, y: 4.312 + 0.154 - g[9] / 2, w: 0.308, h: g[9], rotate: 90,
      line: { color: GREY, width: 0.75, transparency: 50 } });
  });
  // long grey trend arrow rising left to right (dashed lead-in, then solid)
  ln(s, { x: 0.60, y: 4.06, w: 1.20, h: -0.42, line: { color: GREY, width: 3, transparency: 70, dashType: 'lgDash' } });
  ln(s, { x: 1.80, y: 3.64, w: 7.15, h: -2.11, line: { color: GREY, width: 3, transparency: 70, endArrowType: 'triangle' } });
  KOTTER.forEach(function (k) {
    rrect(s, { x: k[2], y: k[3], w: 0.246, h: 4.317 - k[3], adj: 20000, fill: { color: k[5] } });
    rrect(s, { x: k[2] + 0.045, y: k[3] + 0.046, w: 0.156, h: 0.156, adj: 25312, fill: { color: SOFT } });
    txt(s, k[0], { x: k[2] + 0.03, y: k[3] + 0.041, w: 0.178, h: 0.161, fontSize: 7, align: 'center', valign: 'middle' });
    txt(s, k[1], { x: k[2] - 0.256, y: k[4], w: 0.784, h: 0.295, fontSize: 8, color: GREY, align: 'center' });
  });
}

// ── slide 27 · contact ──────────────────────────────────────────────────────
const OFFICES = [
  ['Paris', '10 Avenue Montaigne, 75008 Paris, \nFrance', 'Phone: \n+33-1-345-6789', 0.925],
  ['Berlin', 'Kurfürstendamm 207-208, 10719 Berlin, Germany', 'Phone: +49-30-654-3210', 2.384],
  ['Los Angeles', '789 Sunset Blvd, Los Angeles, CA 90046, USA', 'Phone: +1-555-123-4567', 3.858]
];
function slide27(pptx) {
  const s = page(pptx, 'dark');
  photo(s, { x: 6.128, y: 0.830, w: 2.968, h: 3.943, adj: 8514 });
  stripes(s, { x: -0.005, y: 0, w: 0.474, h: 5.625, fill: { color: GREY, transparency: 70 } });
  txt(s, '07', { x: 0.903, y: 1.084, w: 0.779, h: 0.673, fontSize: 38, color: LIME });
  txt(s, 'Contact', { x: 1.562, y: 1.109, w: 2.723, h: 0.673, fontSize: 38, color: LIGHT, valign: 'bottom' });
  txt(s, 'Include your company name, address, phone number, and email address as basic contact information. Additional details like social ' +
    'media links, feedback forms, and business hours can also be useful. Make sure your contact information is easily accessible for ' +
    'customers to enhance their experience with your company.',
    { x: 0.931, y: 1.723, w: 4.781, h: 0.698, fontSize: 10, color: GREY });
  OFFICES.forEach(function (o) {
    txt(s, o[0], { x: o[3], y: 2.735, w: 0.999, h: 0.214, fontSize: 10, color: LIGHT, valign: 'middle' });
    txt(s, o[1], { x: o[3] - 0.008, y: 2.961, w: 1.304, h: 0.55, fontSize: 10, color: GREY });
    txt(s, '✆', { x: o[3] - 0.02, y: 3.55, w: 0.24, h: 0.24, fontSize: 11, color: LIME, valign: 'middle' });
    txt(s, o[2], { x: o[3] - 0.008, y: 3.80, w: 1.304, h: 0.4, fontSize: 10, color: GREY });
  });
  shieldTick(s, { x: 7.524, y: 2.560, w: 0.217, h: 0.265 });
}

// ── build ───────────────────────────────────────────────────────────────────
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: SW, height: SH });
pptx.layout = 'DECK';
pptx.title = 'Change Management Models';
pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };
SLIDES.forEach(function (build) { build(pptx); });
pptx.writeFile({ fileName: path.join(__dirname, '070d20ac-dc46-4d12-84ed-1575386dc5b5_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
