/**
 * Webinar deck — recreated with pptxgenjs.
 * Run: node 01b6198c-baf8-4af3-9343-a12ca0eddefb_grok_final.js
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const W = 13.333;
const H = 7.5;

/* ------------------------------------------------------------------ palette */
const C = {
  red: 'DF1637',        // accent1
  redDark: 'BE132F',    // accent1, 85% luminance — top of every red gradient
  maroon: '2D0406',     // accent2
  pink: 'F9D0D7',       // accent1 tint
  page: 'EEEEEE',       // slide background (accent5)
  white: 'FFFFFF',
  ink: '262626',        // headline grey
  body: '595959',       // body grey
  rule: 'BFBFBF',       // thin divider rules
  silver: 'E4E4E4',
  quote: 'E4E4E4'
};

const FONT_HEAD = 'Urbanist Medium';
const FONT_BODY = 'Open Sans Light';

/* soft drop shadows used throughout the deck */
const SH_CARD = { type: 'outer', blur: 90, offset: 60, angle: 45, color: '000000', opacity: 0.2 };
const SH_CHIP = { type: 'outer', blur: 25, offset: 3, angle: 45, color: '000000', opacity: 0.2 };
const SH_SOFT = { type: 'outer', blur: 35, offset: 15, angle: 50, color: '000000', opacity: 0.06 };
const SH_RING = { type: 'outer', blur: 50, offset: 10, angle: 90, color: '000000', opacity: 0.15 };

/* ------------------------------------------------------------------ helpers */
const mix = (a, b, t) => {
  const p = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  const v = i => Math.round(p(a, i) + (p(b, i) - p(a, i)) * t).toString(16).padStart(2, '0');
  return (v(0) + v(1) + v(2)).toUpperCase();
};

/** Colour at position t (0..1) across a list of [pos, hex] stops. */
function sample(stops, t) {
  for (let i = 1; i < stops.length; i++) {
    if (t <= stops[i][0] || i === stops.length - 1) {
      const [p0, c0] = stops[i - 1], [p1, c1] = stops[i];
      return mix(c0, c1, p1 === p0 ? 0 : Math.min(1, Math.max(0, (t - p0) / (p1 - p0))));
    }
  }
  return stops[0][1];
}

/**
 * pptxgenjs has no gradient fill, so gradients are painted as a stack of thin
 * solid bands. `angle` is 90 for top-to-bottom; other angles are only used for
 * full-slide washes, where the overspill is clipped by the slide edge.
 */
function gradient(slide, box, stops, angle, steps) {
  const a = angle * Math.PI / 180, ca = Math.cos(a), sa = Math.sin(a);
  const span = Math.abs(box.w * ca) + Math.abs(box.h * sa);   // length of gradient axis
  const wide = Math.abs(box.w * sa) + Math.abs(box.h * ca);   // band length
  const cx = box.x + box.w / 2, cy = box.y + box.h / 2;
  const bw = span / steps;
  for (let i = 0; i < steps; i++) {
    const t = (i + 0.5) / steps;
    const d = (t - 0.5) * span;
    slide.addShape('rect', {
      x: cx + d * ca - bw * 0.75, y: cy + d * sa - wide / 2,
      w: bw * 1.5, h: wide, rotate: angle,
      fill: { color: sample(stops, t) }, line: { type: 'none' }
    });
  }
}

/** The deck's signature panel fill: bright red at the top fading to near-black. */
const panel = (slide, box, steps) =>
  gradient(slide, box, [[0, C.redDark], [1, C.maroon]], 90, steps || 40);

/** Rotated square with the same top-to-bottom red gradient (the "diamond" motif). */
function diamond(slide, cx, cy, side, rot, steps) {
  steps = steps || 26;
  const a = rot * Math.PI / 180, bh = side / steps;
  for (let i = 0; i < steps; i++) {
    const d = ((i + 0.5) / steps - 0.5) * side;
    slide.addShape('rect', {
      x: cx - d * Math.sin(a) - side / 2, y: cy + d * Math.cos(a) - bh * 0.55,
      w: side, h: bh * 1.1, rotate: rot,
      fill: { color: mix(C.redDark, C.maroon, (i + 0.5) / steps) }, line: { type: 'none' }
    });
  }
}

/* ----------------------------------------------------------- text shorthands */
const text = (slide, str, o) => slide.addText(str, Object.assign({
  fontFace: FONT_BODY, fontSize: 18, color: C.body, valign: 'top', margin: [7.2, 7.2, 3.6, 3.6]
}, o));

/** 48pt bold headline, 0.8 line spacing — used on nearly every slide. */
const title = (slide, str, x, y, w, h, color) =>
  text(slide, str, { x, y, w, h, fontFace: FONT_HEAD, fontSize: 48, bold: true, lineSpacingMultiple: 0.8, color: color || C.ink });

/** 14pt grey paragraph, 1.2 line spacing. */
const para = (slide, str, x, y, w, h, o) =>
  text(slide, str, Object.assign({ x, y, w, h, fontSize: 14, lineSpacingMultiple: 1.2 }, o));

/** Thin vertical divider rule. */
const vrule = (slide, x, y, h, o) => slide.addShape('line', Object.assign({
  x, y, w: 0, h, line: { color: C.rule, width: 1.5, transparency: 50 }
}, o));

/** Round red badge holding a white chevron — the deck's bullet marker. */
function badge(slide, x, y, size) {
  const d = size || 0.343;
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: 'B0122B' }, line: { type: 'none' } });
  slide.addText('\u203A', {
    x, y: y - 0.02, w: d, h: d, align: 'center', valign: 'middle',
    fontFace: FONT_HEAD, fontSize: 14, bold: true, color: C.white, margin: 0
  });
}

/** Red disc with a white tick — the check bullet on slides 13 and 17. */
function tick(slide, x, y, label, size) {
  const d = 0.303;
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: C.red }, line: { type: 'none' } });
  slide.addText('\u2713', {
    x, y, w: d, h: d, align: 'center', valign: 'middle',
    fontFace: 'Open Sans', fontSize: 11, bold: true, color: C.white, margin: 0
  });
  para(slide, label, x + 0.429, y + 0.009, 4.164, 0.364, { fontSize: size || 14 });
}

/** Big number over a small caption, the layout used for every KPI in the deck. */
function stat(slide, o) {
  text(slide, o.value, {
    x: o.x, y: o.y, w: o.vw, h: o.vh || 0.572, fontFace: FONT_HEAD, wrap: o.nowrap !== true,
    fontSize: o.size || 28, bold: true, color: o.color || C.ink, lineSpacingMultiple: 1.0
  });
  text(slide, o.label, {
    x: o.x, y: o.ly, w: o.lw, h: 0.326, fontSize: o.labelSize || 12,
    color: o.labelColor || C.body, fontFace: o.labelFont || FONT_BODY, lineSpacingMultiple: 1.2
  });
}

/** White card with the deck's oversized soft shadow. */
const card = (slide, x, y, w, h, o) => slide.addShape('rect', Object.assign({
  x, y, w, h, fill: { color: C.white }, line: { type: 'none' }, shadow: SH_CARD
}, o));

/**
 * Phone mock-up standing in for the product screenshots. `screen` paints the
 * stock wallpaper; without it the handset is an empty outline as on slide 11.
 */
function phone(slide, x, y, w, h, screen) {
  const r = w * 0.155, m = w * 0.05;
  if (screen) {
    slide.addShape('roundRect', {
      x, y, w, h, rectRadius: r, fill: { color: '111114' }, line: { type: 'none' }, shadow: SH_SOFT
    });
    const bands = 26, bh = (h - 2 * m) / bands;
    for (let i = 0; i < bands; i++) {
      slide.addShape('rect', {
        x: x + m, y: y + m + i * bh, w: w - 2 * m, h: bh * 1.2,
        fill: { color: sample([[0, 'E72B83'], [0.35, 'EA5764'], [0.72, 'F7D3B8'], [1, '9FB6DC']], (i + 0.5) / bands) },
        line: { type: 'none' }
      });
    }
    // thick ring in the frame colour rounds off the square band corners
    slide.addShape('roundRect', {
      x: x + m / 2, y: y + m / 2, w: w - m, h: h - m, rectRadius: r * 0.9,
      fill: { type: 'none' }, line: { color: '111114', width: m * 72 }
    });
  } else {
    // Empty handset — the source PNG has a see-through screen, so the diamond
    // motif behind it stays visible.
    slide.addShape('roundRect', { x, y, w, h, rectRadius: r, fill: { type: 'none' }, line: { color: '3A3A3E', width: 5 } });
    slide.addShape('roundRect', {
      x: x + m, y: y + m, w: w - 2 * m, h: h - 2 * m, rectRadius: r * 0.85,
      fill: { type: 'none' }, line: { color: 'C8C8CC', width: 1.5 }
    });
  }
  slide.addShape('roundRect', {
    x: x + w * 0.30, y: y + h * 0.010, w: w * 0.40, h: h * 0.026,
    rectRadius: h * 0.013, fill: { color: '111114' }, line: { type: 'none' }
  });
}

/* --------------------------------------------------------- small vector icons */
/** Calendar glyph (slide 15 timeline markers). */
function iconCalendar(slide, x, y, d, color) {
  slide.addShape('rect', { x, y: y + d * 0.14, w: d, h: d * 0.86, fill: { type: 'none' }, line: { color, width: 1.5 } });
  slide.addShape('rect', { x, y: y + d * 0.14, w: d, h: d * 0.26, fill: { color }, line: { type: 'none' } });
  [0.28, 0.68].forEach(f => slide.addShape('rect', {
    x: x + d * f, y, w: d * 0.06, h: d * 0.2, fill: { color }, line: { type: 'none' }
  }));
}

/** Pie glyph (slides 11 and 19). */
function iconPie(slide, x, y, d, color) {
  slide.addShape('pie', { x, y, w: d, h: d, angleRange: [270, 180], fill: { color }, line: { type: 'none' } });
  slide.addShape('pie', { x, y, w: d, h: d, angleRange: [190, 260], fill: { color }, line: { type: 'none' } });
}

/** Three interlocking circles (slide 19). */
function iconVenn(slide, x, y, d, color) {
  const r = d * 0.56;
  [[d * 0.22, 0], [0, d * 0.44], [d * 0.44, d * 0.44]].forEach(([dx, dy]) =>
    slide.addShape('ellipse', { x: x + dx, y: y + dy, w: r, h: r, fill: { color }, line: { color: C.white, width: 0.75 } }));
}

/** Flip-chart glyph on the red tile of slide 11. */
function iconBoard(slide, x, y, d) {
  slide.addShape('rect', { x, y: y + d * 0.1, w: d, h: d * 0.55, fill: { type: 'none' }, line: { color: C.white, width: 1.75 } });
  slide.addShape('line', { x, y: y + d * 0.1, w: d, h: 0, line: { color: C.white, width: 2.5 } });
  iconPie(slide, x + d * 0.26, y + d * 0.2, d * 0.34, C.white);
  slide.addShape('line', { x: x + d * 0.5, y: y + d * 0.65, w: 0, h: d * 0.3, line: { color: C.white, width: 1.5 } });
  slide.addShape('line', { x: x + d * 0.18, y: y + d * 0.65, w: d * 0.32, h: d * 0.3, line: { color: C.white, width: 1.5 } });
  slide.addShape('line', { x: x + d * 0.5, y: y + d * 0.65, w: d * 0.32, h: d * 0.3, line: { color: C.white, width: 1.5 }, flipV: true });
}

/* ------------------------------------------------------------------- slides */
/** Every slide starts on the deck's light-grey page colour. */
const newSlide = pptx => {
  const s = pptx.addSlide();
  s.background = { color: C.page };
  return s;
};


// 1 — Title
function slide01(pptx) {
  const s = newSlide(pptx);
  gradient(s, { x: 0, y: 0, w: W, h: H }, [[0, C.redDark], [1, C.maroon]], 44, 56);
  text(s, 'Speaker Webinar', {
    x: 0.922, y: 2.367, w: 7.102, h: 2.767, fontFace: FONT_HEAD, fontSize: 88,
    bold: true, color: C.white, lineSpacingMultiple: 0.9
  });
  panel(s, { x: 7.986, y: 3.173, w: 2.673, h: 4.327 });
  panel(s, { x: 10.660, y: 0, w: 2.673, h: 3.173 }, 30);
  [7.986, 10.660].forEach(x => s.addShape('line', {
    x, y: 0, w: 0, h: H, line: { color: C.red, width: 0.75, transparency: 40 }
  }));
  text(s, 'Presentation Template', { x: 8.208, y: 3.355, w: 1.938, h: 0.707, color: C.white });
  s.addShape('line', { x: 8.024, y: 3.173, w: 5.309, h: 0, line: { color: C.red, width: 0.75, transparency: 40 } });
  text(s, 'www.yourwebsite.com', { x: 8.208, y: 6.514, w: 1.938, h: 0.303, fontSize: 12, color: C.white });
}

// 2 — Welcome
function slide02(pptx) {
  const s = newSlide(pptx);
  title(s, 'Welcome to the Webinar', 0.893, 1.131, 7.810, 0.754);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed', 8.618, 3.033, 3.965, 0.929);
  text(s, 'An Introduction to Today\u2019s Theme & Speaker', {
    x: 5.128, y: 5.095, w: 3.365, h: 0.707, fontFace: FONT_HEAD, color: C.red, lineSpacingMultiple: 1.0
  });
  vrule(s, 8.503, 5.015, 0.866);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', 9.563, 5.125, 2.791, 0.646);
  badge(s, 8.780, 5.276);
}

// 3 — Ideal composition
function slide03(pptx) {
  const s = newSlide(pptx);
  card(s, 0.978, 4.073, 6.544, 2.007);
  title(s, 'Ideal Composition to Convey a Message', 0.978, 1.420, 6.544, 1.400);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar', 0.978, 3.001, 6.086, 0.646);
  [{ x: 10.861, v: '98+', vw: 1.333, l: 'Partner business', lw: 1.333 },
   { x: 8.699, v: '150+', vw: 1.651, l: 'Award achieved', lw: 1.880 }].forEach(k =>
    stat(s, {
      x: k.x, y: 4.360, vw: k.vw, vh: 0.805, value: k.v, size: 40, color: C.red,
      ly: 5.087, lw: k.lw, label: k.l, labelSize: 18, labelColor: C.ink, labelFont: FONT_HEAD
    }));
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor', 4.250, 4.612, 3.091, 0.929);
  vrule(s, 3.789, 4.397, 1.359);
  badge(s, 3.059, 4.905);
}

// 4 — Visualization that blends with branding
function slide04(pptx) {
  const s = newSlide(pptx);
  card(s, 3.782, 0.865, 8.121, 3.354);
  title(s, 'Visualization that Blends with Branding', 4.659, 1.451, 6.544, 1.400);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar', 4.659, 2.986, 6.086, 0.646);
  text(s, 'Session Flow and Key Segments', {
    x: 2.893, y: 5.237, w: 4.094, h: 0.404, fontFace: FONT_HEAD, color: C.red, lineSpacingMultiple: 1.0
  });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor', 2.893, 5.641, 3.958, 0.646);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed', 8.139, 5.298, 3.965, 0.929);
  vrule(s, 7.202, 5.097, 1.330);
  badge(s, 7.445, 5.591);
}

// 5 — Key questions answered
function slide05(pptx) {
  const s = newSlide(pptx);
  gradient(s, { x: 0, y: 0, w: W, h: H }, [[0, C.red], [0.47, C.red], [1, 'EAC1C8']], -31, 48);
  title(s, 'Key Questions Answered', 1.096, 2.006, 4.837, 1.400, C.white);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, ', 1.096, 3.768, 4.837, 0.929, { color: C.white });
  ['$ 40 Million', '20,000', '30% Margin', '$ 5 Million'].forEach((v, i) => {
    const x = 1.096 + i * 2.660;
    stat(s, {
      x, y: 5.422, vw: 2.396, vh: 0.505, value: v, size: 24, color: C.white,
      ly: 5.867, lw: 1.932, label: 'Lorem ipsum dolor ', labelColor: C.white
    });
    if (i) vrule(s, x - 0.132, 5.377, 0.866, { line: { color: C.white, width: 1.5, transparency: 50 } });
  });
}

// 6 — Promises of unbeatable prices
function slide06(pptx) {
  const s = newSlide(pptx);
  title(s, 'Promises of unbeatable prices and hidden treasures', 1.030, 1.288, 5.637, 2.686);
  panel(s, { x: 8.115, y: 0, w: 4.049, h: H });
  card(s, 8.545, 1.186, 3.986, 1.314);
  const items = [
    { y: 1.330, t: 'Mapping actual versus forecasted market sales.', c: C.ink },
    { y: 3.245, t: 'The diagnostic determined problem advantage.', c: 'C9A0A6' },
    { y: 5.161, t: 'Improved the company\u2019s planning and execution.', c: 'C9A0A6' }
  ];
  items.forEach(it => text(s, it.t, {
    x: 8.863, y: it.y, w: 2.736, h: 1.010, fontFace: FONT_HEAD, color: it.c, lineSpacingMultiple: 1.0
  }));
  [2.792, 4.708].forEach(y => s.addShape('line', {
    x: 8.741, y, w: 2.682, h: 0, line: { color: 'B96E7B', width: 1 }
  }));
  badge(s, 8.365, 1.663);
  para(s, 'Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit.', 4.548, 5.284, 3.002, 0.928);
  text(s, 'A unique note in the melody of sales', {
    x: 4.548, y: 4.501, w: 2.827, h: 0.707, fontFace: FONT_HEAD, color: C.red, lineSpacingMultiple: 1.0
  });
}

// 7 — Shape an enterprise's journey
function slide07(pptx) {
  const s = newSlide(pptx);
  s.addShape('rect', { x: 0, y: 1.043, w: 6.958, h: 4.349, fill: { color: C.ink, transparency: 35 }, line: { type: 'none' } });
  title(s, "Shape an enterprise's journey", 1.864, 2.021, 4.793, 2.040, C.white);
  card(s, 2.036, 4.737, 4.521, 1.661);
  const note = 'Nam libero tempore, cum soluta nobis est eligendi optio cumque nihil impedit quo minus id quod maxime.';
  para(s, note, 2.622, 5.089, 3.473, 0.923);
  badge(s, 1.864, 5.395);
  para(s, note, 8.703, 5.103, 3.473, 0.923);
  badge(s, 7.919, 5.400);
  vrule(s, 7.611, 5.134, 0.866);
  [{ x: 7.919, y: 1.398, v: '15+', vw: 1.651, l: 'Years experience', lw: 1.880 },
   { x: 7.919, y: 2.930, v: '150+', vw: 1.651, l: 'Award achieved', lw: 1.880 },
   { x: 10.287, y: 1.398, v: '98+', vw: 1.651, l: 'Partner business', lw: 1.880 },
   { x: 10.287, y: 2.930, v: '72k+', vw: 1.651, l: 'Client satisfied', lw: 1.651 }].forEach(k =>
    stat(s, {
      x: k.x, y: k.y, vw: k.vw, vh: 0.79, value: k.v, size: 40, color: C.red,
      ly: k.y + 0.727, lw: k.lw, label: k.l, labelSize: 18, labelColor: C.ink, labelFont: FONT_HEAD
    }));
}

// 8 — From career to future vision
function slide08(pptx) {
  const s = newSlide(pptx);
  panel(s, { x: 6.667, y: 0, w: 6.667, h: H });
  diamond(s, 1.909 + 1.4585, 3.369 + 1.4585, 2.917, 43.7);
  title(s, 'From Career to Future Vision', 7.398, 1.904, 4.841, 1.400, C.white);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies', 7.398, 3.483, 4.878, 0.929, { color: C.white });
  text(s, 'Mark Davis', { x: 9.573, y: 4.889, w: 1.803, h: 0.404, fontFace: FONT_HEAD, color: C.white });
  text(s, 'Your Position', { x: 9.573, y: 5.247, w: 1.603, h: 0.286, fontFace: FONT_HEAD, fontSize: 11, italic: true, color: C.white });
  text(s, 'Who Are We Talking To?', { x: 7.398, y: 4.889, w: 1.803, h: 0.707, fontFace: FONT_HEAD, color: C.white });
  s.addShape('line', { x: 9.387, y: 4.889, w: 0, h: 0.707, line: { color: C.white, width: 1 } });
}

// 9 — The great people behind the webinar
function slide09(pptx) {
  const s = newSlide(pptx);
  diamond(s, 1.655 + 0.7615, 4.729 + 0.7615, 1.523, 45.3, 18);
  diamond(s, 3.724 + 0.7615, 2.234 + 0.7615, 1.523, 45.3, 18);
  [{ x: 4.260, y: 4.506, n: 'Maria Aulia' }, { x: 1.819, y: 1.578, n: 'Ken Egbert' }].forEach(p => {
    text(s, p.n, { x: p.x, y: p.y, w: 1.803, h: 0.404, fontFace: FONT_HEAD, color: C.ink });
    text(s, 'Your Position', { x: p.x, y: p.y + 0.358, w: 1.603, h: 0.286, fontFace: FONT_HEAD, fontSize: 11, italic: true, color: C.ink });
  });
  title(s, 'The Great People Behind the Webinar', 6.738, 1.357, 4.864, 2.046);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna', 6.738, 3.674, 5.584, 0.646);
  text(s, 'An Award for Collaboration and Dedication', { x: 9.530, y: 4.928, w: 2.570, h: 1.010, fontFace: FONT_HEAD, color: C.ink });
  [[8.745, 5.261], [3.765, 4.656], [1.340, 1.728]].forEach(([x, y]) => badge(s, x, y));
}

// 10 — Break section
function slide10(pptx) {
  const s = newSlide(pptx);
  gradient(s, { x: 0, y: 0, w: W, h: H }, [[0, C.redDark], [1, C.maroon]], 44, 56);
  text(s, 'Break Section', {
    x: 1.386, y: 0.983, w: 4.864, h: 2.767, fontFace: FONT_HEAD, fontSize: 88,
    bold: true, color: C.white, lineSpacingMultiple: 0.9
  });
  text(s, 'Use This Moment to Absorb and Reflect on the Beginning Material', {
    x: 7.630, y: 1.686, w: 4.278, h: 0.646, fontFace: FONT_HEAD, color: C.white, lineSpacingMultiple: 0.9
  });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce', 7.069, 2.401, 4.878, 0.646, { color: C.white });
  badge(s, 7.157, 1.837);
}

// 11 — Visualization of real implementation
function slide11(pptx) {
  const s = newSlide(pptx);
  diamond(s, 8.673 + 1.5, 3.122 + 1.5, 3.0, 45.2, 30);
  title(s, 'Visualization of Real Implementation', 1.323, 1.892, 6.445, 1.400);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 1.323, 3.455, 5.469, 0.929);
  text(s, 'Live Overview of How to Implement', {
    x: 2.251, y: 4.883, w: 2.851, h: 0.707, fontFace: FONT_HEAD, color: C.red, lineSpacingMultiple: 1.0
  });
  panel(s, { x: 1.323, y: 4.864, w: 0.744, h: 0.744 }, 14);
  iconBoard(s, 1.421 + 0.09, 4.962 + 0.09, 0.548 - 0.18);
  phone(s, 8.618, 1.046, 3.108, 5.682, false);
}

// 12 — Product screenshot
function slide12(pptx) {
  const s = newSlide(pptx);
  title(s, 'Product Screenshot or Material Discussed', 1.155, 1.488, 6.876, 1.400);
  phone(s, 5.563, 3.335, 3.224, 6.275, true);
  phone(s, 8.952, 0.757, 3.224, 6.275, true);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna', 1.155, 3.335, 4.001, 0.929);
  optionPair(s, 4.788);
}

/** "150+ / 72k+  Option Here" duo used on slides 12 and 13. */
function optionPair(s, y) {
  stat(s, { x: 1.272, y, vw: 1.297, value: '150+', color: C.ink, ly: y + 0.545, lw: 1.297, label: 'Option Here' });
  stat(s, { x: 2.947, y, vw: 1.250, value: '72k+', color: C.red, ly: y + 0.545, lw: 1.139, label: 'Option Here' });
  vrule(s, 2.758, y - 0.048, 0.968);
}

// 13 — Facts that support today's presentation (bar chart)
function slide13(pptx) {
  const s = newSlide(pptx);
  card(s, 7.173, 0.958, 4.889, 5.583);
  s.addChart(pptx.ChartType.bar, [
    { name: 'Series 1', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'], values: [2, 2, 3, 5] }
  ], {
    x: 7.551, y: 1.120, w: 4.132, h: 5.260,
    barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 219, barOverlapPct: -27,
    chartColors: [C.red, C.maroon, C.white],
    showTitle: true, title: 'Chart Title', titleFontSize: 18, titleColor: C.body, titleFontFace: FONT_BODY,
    showLegend: false, showValue: false,
    catAxisLabelColor: C.body, catAxisLabelFontFace: FONT_BODY, catAxisLabelFontSize: 12, catAxisLineShow: true,
    valAxisLabelColor: C.body, valAxisLabelFontFace: FONT_BODY, valAxisLabelFontSize: 12, valAxisLineShow: false,
    valGridLine: { color: 'D9D9D9', size: 0.75 }, catGridLine: { style: 'none' }
  });
  title(s, "Facts That Support Today's Presentation", 1.272, 1.182, 5.790, 2.046);
  optionPair(s, 4.996);
  tick(s, 1.272, 3.567, 'Lorem ipsum dolor sit amet, consectetuer', 12);
  tick(s, 1.272, 4.183, 'Lorem ipsum dolor sit amet, consectetuer', 12);
}

// 14 — Seeing change and impact visually (pie chart)
function slide14(pptx) {
  const s = newSlide(pptx);
  // the reference uses a 3-D pie; a flat pie plus a shaded skirt underneath
  // gives the same silhouette with a native chart
  s.addShape('ellipse', { x: 6.48, y: 3.10, w: 5.45, h: 2.35, fill: { color: '9E1029' }, line: { type: 'none' } });
  s.addChart(pptx.ChartType.pie, [
    { name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [8.2, 3.2] }
  ], {
    x: 6.030, y: 0.92, w: 6.386, h: 4.30,
    chartColors: [C.red, C.maroon], dataBorder: { pt: 2, color: C.white },
    showTitle: true, title: 'Sales', titleFontSize: 18, titleColor: C.body, titleFontFace: FONT_BODY,
    showLegend: false, showValue: false, firstSliceAng: 0
  });
  text(s, '30%', { x: 7.226, y: 2.493, w: 1.442, h: 0.439, align: 'center', fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.white, lineSpacingMultiple: 0.8 });
  text(s, '70%', { x: 9.352, y: 3.300, w: 1.442, h: 0.439, align: 'center', fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.white, lineSpacingMultiple: 0.8 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa', 6.489, 5.786, 5.469, 0.646, { align: 'center' });
  title(s, 'Seeing Change and Impact Visually', 1.272, 1.471, 5.218, 2.046);
  profitCards(s, 4.790);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor', 1.272, 3.741, 4.001, 0.646);
}

/** The paired 53% / 78% chips used on slides 14 and 17. */
function profitCards(s, y) {
  [{ x: 3.530, fill: C.red, v: '78%', vw: 1.057, disc: C.silver, tri: C.red, up: true },
   { x: 1.259, fill: C.maroon, v: '53%', vw: 1.036, disc: C.white, tri: C.maroon, up: false }].forEach(k => {
    s.addShape('roundRect', { x: k.x, y, w: 2.046, h: 1.239, rectRadius: 0.062, fill: { color: k.fill }, line: { type: 'none' }, shadow: SH_CHIP });
    text(s, k.v, { x: k.x + 0.111, y: y + 0.108, w: k.vw, h: 0.640, fontFace: FONT_HEAD, fontSize: 32, color: C.white, wrap: false });
    text(s, 'Profit Here', { x: k.x + 0.111, y: y + 0.769, w: 1.281, h: 0.308, fontSize: 12, color: C.white });
    const cx = k.x + 1.590;
    s.addShape('ellipse', { x: cx, y: y + 0.108, w: 0.328, h: 0.328, fill: { color: k.disc }, line: { type: 'none' } });
    s.addShape('triangle', { x: cx + 0.084, y: y + 0.184, w: 0.161, h: 0.133, fill: { color: k.tri }, line: { type: 'none' }, flipV: !k.up });
  });
}

// 15 — The speaker's career journey
function slide15(pptx) {
  const s = newSlide(pptx);
  title(s, "The Speaker's Career Journey", 1.264, 1.223, 5.218, 1.400);
  ['2020', '2022', '2023', '2024'].forEach((year, i) => {
    const x = 1.264 + i * 2.786, dark = i % 2 === 1;
    s.addShape('roundRect', {
      x, y: 3.618, w: 2.448, h: 2.659, rectRadius: 0.203,
      fill: { color: dark ? C.maroon : C.red }, line: { type: 'none' }
    });
    s.addShape('ellipse', {
      x: x + 0.751, y: 3.139, w: 0.945, h: 0.945,
      fill: { color: dark ? C.maroon : C.white }, line: { type: 'none' }, shadow: SH_RING
    });
    text(s, year, { x: x + 0.348, y: 4.366, w: 1.751, h: 0.640, align: 'center', fontFace: FONT_HEAD, fontSize: 32, bold: true, color: C.white });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', {
      x: x + 0.174, y: 5.063, w: 2.100, h: 0.811, align: 'center', fontSize: 12, color: C.white, lineSpacingMultiple: 1.2
    });
    s.addShape('ellipse', {
      x: x + 0.842, y: 3.229, w: 0.764, h: 0.764,
      fill: { color: dark ? C.white : C.red }, line: { type: 'none' }
    });
    iconCalendar(s, x + 1.017, 3.404, 0.413, dark ? C.maroon : C.white);
  });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et', 6.836, 1.458, 4.886, 0.929);
}

// 16 — What will be discussed
function slide16(pptx) {
  const s = newSlide(pptx);
  supportAgentArt(s);
  title(s, 'What Will Be Discussed', 1.264, 1.425, 4.444, 1.400);
  [{ x: 1.264, v: '$56', vw: 0.866, c: C.ink, lw: 1.224 },
   { x: 2.928, v: '86%', vw: 0.882, c: C.red, lw: 1.222 },
   { x: 4.492, v: '123+', vw: 0.952, c: C.ink, lw: 1.216 }].forEach((k, i) => {
    stat(s, { x: k.x, y: 4.400 + (i === 2 ? -0.007 : 0.008), vw: k.vw, value: k.v, color: k.c, nowrap: true, ly: 4.983 + (i === 2 ? 0.015 : 0), lw: k.lw, label: 'Option Here' });
    if (i) vrule(s, [0, 2.708, 4.272][i], 4.347, 1.0);
  });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis', 1.323, 3.177, 5.469, 0.929);
  panel(s, { x: 0, y: 5.990, w: W, h: 1.510 }, 18);
}

/** Simplified stand-in for the head-set agent illustration (right of slide 16). */
function supportAgentArt(s) {
  // pale cloud of chat/commerce icons behind the figure
  [[8.94, 2.46], [8.83, 1.61], [9.29, 3.16], [8.07, 3.68], [11.24, 2.95], [8.67, 3.24],
   [11.72, 4.26], [11.93, 2.93], [8.20, 4.34], [11.44, 3.58], [11.30, 4.98], [10.63, 3.53],
   [10.80, 4.16], [9.82, 1.15], [8.19, 5.00], [9.07, 4.11], [10.49, 2.79], [8.03, 2.90],
   [10.79, 1.42], [9.91, 3.45], [10.53, 5.09], [9.86, 4.37], [9.90, 2.01], [9.73, 2.74],
   [10.55, 2.20], [11.54, 2.05], [9.17, 4.96], [8.22, 2.20], [8.97, 3.62]]
    .forEach(([x, y], i) => s.addShape(['roundRect', 'ellipse', 'roundRect', 'wedgeRoundRectCallout'][i % 4], {
      x, y, w: 0.70, h: 0.62, rectRadius: 0.14, fill: { color: 'DEDEDE' }, line: { type: 'none' }
    }));
  s.addShape('ellipse', { x: 8.995, y: 1.317, w: 2.175, h: 3.163, fill: { color: '5C381C' }, line: { type: 'none' } });   // hair silhouette
  s.addShape('rect', { x: 8.995, y: 2.60, w: 2.175, h: 2.70, fill: { color: '84593D' }, line: { type: 'none' } });        // hair down to shoulders
  s.addShape('rect', { x: 9.75, y: 3.60, w: 0.42, h: 0.90, fill: { color: 'FFF1D2' }, line: { type: 'none' } });          // neck
  s.addShape('rect', { x: 10.17, y: 3.60, w: 0.36, h: 0.90, fill: { color: 'E9DCC3' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 8.496, y: 4.087, w: 3.348, h: 2.10, rectRadius: 0.42, fill: { color: C.red }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 10.170, y: 4.087, w: 1.673, h: 2.10, rectRadius: 0.42, fill: { color: 'A71029' }, line: { type: 'none' } });
  s.addShape('rect', { x: 8.734, y: 4.514, w: 0.575, h: 1.574, fill: { color: 'A71029' }, line: { type: 'none' } });      // sleeve shading
  s.addShape('rect', { x: 11.032, y: 4.514, w: 0.575, h: 1.574, fill: { color: '700B1C' }, line: { type: 'none' } });
  s.addShape('triangle', { x: 9.60, y: 4.006, w: 0.57, h: 0.90, flipV: true, fill: { color: 'D9D9D9' }, line: { type: 'none' } }); // collar
  s.addShape('triangle', { x: 10.17, y: 4.006, w: 0.57, h: 0.90, flipV: true, fill: { color: 'F2F2F2' }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 9.186, y: 1.502, w: 1.968, h: 2.351, fill: { color: 'FFF1D2' }, line: { type: 'none' } });   // face
  s.addShape('pie', { x: 9.186, y: 1.502, w: 1.968, h: 2.351, angleRange: [270, 90], fill: { color: 'E9DCC3' }, line: { type: 'none' } });
  s.addShape('pie', { x: 9.240, y: 1.317, w: 1.929, h: 1.40, angleRange: [180, 360], fill: { color: '754B2E' }, line: { type: 'none' } }); // fringe
  s.addShape('rect', { x: 9.240, y: 1.90, w: 1.929, h: 0.32, fill: { color: '754B2E' }, line: { type: 'none' } });
  s.addShape('rect', { x: 9.044, y: 2.60, w: 0.434, h: 2.360, fill: { color: '754B2E' }, line: { type: 'none' } });       // strands beside the face
  s.addShape('rect', { x: 10.80, y: 2.60, w: 0.370, h: 2.360, fill: { color: '754B2E' }, line: { type: 'none' } });
  s.addShape('arc', { x: 9.15, y: 1.25, w: 2.05, h: 1.60, angleRange: [185, 355], line: { color: '262626', width: 5 } }); // head band
  s.addShape('roundRect', { x: 9.118, y: 2.377, w: 0.286, h: 0.565, rectRadius: 0.12, fill: { color: '7F7F7F' }, line: { type: 'none' } }); // ear cup
  s.addShape('line', { x: 9.265, y: 2.90, w: 0.79, h: 0.62, line: { color: '404040', width: 4 } });                        // mic boom
  s.addShape('rect', { x: 10.025, y: 3.480, w: 0.281, h: 0.135, fill: { color: '262626' }, line: { type: 'none' } });      // mic
}

// 17 — Who joined today
function slide17(pptx) {
  const s = newSlide(pptx);
  presenterArt(s);
  title(s, 'Who Joined Today', 1.340, 1.615, 5.893, 0.754);
  text(s, 'Trends and Opportunities in Our Industry', { x: 1.340, y: 2.760, w: 4.610, h: 0.370, fontFace: FONT_HEAD, fontSize: 16, color: C.ink });
  tick(s, 1.449, 3.313, 'Lorem ipsum dolor sit amet, consectetuer');
  tick(s, 1.449, 3.838, 'Aenean massa. Cum sociis natoque');
  profitCards(s, 4.646);
}

/** Simplified stand-in for the presenter + audience illustration (right of slide 17). */
function presenterArt(s) {
  // flip chart
  s.addShape('rect', { x: 9.287, y: 1.249, w: 2.257, h: 2.622, fill: { color: 'C5D0D9' }, line: { type: 'none' } });
  s.addShape('rect', { x: 9.337, y: 1.698, w: 2.156, h: 2.13, fill: { color: C.white }, line: { type: 'none' } });
  s.addShape('rect', { x: 9.337, y: 1.299, w: 2.156, h: 0.399, fill: { color: 'C5D0D9' }, line: { type: 'none' } });
  [[9.493, 2.005, 0.159], [9.493, 2.269, 0.024], [9.493, 2.332, 0.024], [9.493, 2.393, 0.032],
   [9.493, 2.457, 0.032], [9.493, 2.520, 0.029], [10.452, 2.914, 0.151], [10.452, 3.170, 0.032],
   [10.452, 3.234, 0.032], [10.452, 3.302, 0.024], [10.452, 3.363, 0.026], [10.452, 3.427, 0.032]]
    .forEach(([x, y, h]) => s.addShape('rect', { x, y, w: 0.835, h, fill: { color: 'CCCCCC' }, line: { type: 'none' } }));
  [[10.452, 2.338, 0.370, 'B7D5F0'], [10.629, 2.192, 0.515, '99BDDC'], [10.799, 2.356, 0.351, '7BA4C9'],
   [10.968, 2.174, 0.534, '5D8CB5'], [11.137, 1.854, 0.854, '3F73A1']]
    .forEach(([x, y, h, c]) => s.addShape('rect', { x, y, w: 0.15, h, fill: { color: c }, line: { type: 'none' } }));
  s.addShape('rect', { x: 10.452, y: 2.708, w: 0.835, h: 0.042, fill: { color: 'FC8D25' }, line: { type: 'none' } });
  s.addShape('ellipse', { x: 9.530, y: 2.951, w: 0.708, h: 0.708, fill: { color: 'B7D5F0' }, line: { type: 'none' } });
  s.addShape('pie', { x: 9.53, y: 2.951, w: 0.708, h: 0.708, angleRange: [0, 100], fill: { color: '5D8CB5' }, line: { type: 'none' } });
  s.addShape('pie', { x: 9.53, y: 2.951, w: 0.708, h: 0.708, angleRange: [270, 355], fill: { color: '7BA4C9' }, line: { type: 'none' } });
  s.addShape('rect', { x: 11.155, y: 3.471, w: 0.288, h: 0.307, fill: { color: 'E8E7E5' }, line: { type: 'none' } });

  // presenter
  s.addShape('rect', { x: 7.857, y: 5.192, w: 3.687, h: 0.388, fill: { color: 'C5D0D9' }, line: { type: 'none' } }); // floor
  s.addShape('rect', { x: 8.383, y: 3.765, w: 0.288, h: 1.522, fill: { color: '27557D' }, line: { type: 'none' } }); // legs
  s.addShape('rect', { x: 8.716, y: 3.453, w: 0.288, h: 1.834, fill: { color: '27557D' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: 8.388, y: 5.192, w: 0.370, h: 0.320, rectRadius: 0.09, fill: { color: '57291E' }, line: { type: 'none' } }); // shoes
  s.addShape('roundRect', { x: 8.716, y: 5.192, w: 0.425, h: 0.301, rectRadius: 0.09, fill: { color: '57291E' }, line: { type: 'none' } });
  s.addShape('rect', { x: 8.584, y: 2.457, w: 0.219, h: 0.182, fill: { color: 'D1A886' }, line: { type: 'none' } }); // neck
  s.addShape('rect', { x: 8.190, y: 2.520, w: 0.864, h: 1.446, fill: { color: '3F73A1' }, line: { type: 'none' } }); // jacket
  s.addShape('rect', { x: 8.428, y: 2.531, w: 0.532, h: 0.702, fill: { color: '27557D' }, line: { type: 'none' } }); // jacket front
  s.addShape('triangle', { x: 8.497, y: 2.531, w: 0.394, h: 0.560, flipV: true, fill: { color: C.white }, line: { type: 'none' } }); // shirt V
  s.addShape('rect', { x: 8.639, y: 2.813, w: 0.114, h: 0.420, fill: { color: 'FC8D25' }, line: { type: 'none' } }); // tie
  s.addShape('rect', { x: 8.671, y: 3.233, w: 0.050, h: 0.677, fill: { color: '27557D' }, line: { type: 'none' } }); // jacket seam
  s.addShape('rect', { x: 8.830, y: 2.388, w: 1.039, h: 0.412, fill: { color: '27557D' }, line: { type: 'none' } }); // raised arm
  s.addShape('rect', { x: 7.963, y: 2.700, w: 0.465, h: 1.142, fill: { color: '27557D' }, line: { type: 'none' } }); // hanging arm
  s.addShape('roundRect', { x: 7.939, y: 3.701, w: 0.230, h: 0.357, rectRadius: 0.09, fill: { color: 'F2CAA9' }, line: { type: 'none' } }); // hand
  s.addShape('ellipse', { x: 9.789, y: 2.388, w: 0.124, h: 0.225, fill: { color: C.white }, line: { type: 'none' } }); // cuff
  s.addShape('line', { x: 9.80, y: 2.50, w: 1.05, h: 0.62, flipV: true, line: { color: '57291E', width: 3.5 } });      // pointer
  s.addShape('ellipse', { x: 8.264, y: 1.807, w: 0.758, h: 0.801, fill: { color: 'F2CAA9' }, line: { type: 'none' } });// head
  s.addShape('pie', { x: 8.30, y: 1.79, w: 0.70, h: 0.62, angleRange: [178, 2], fill: { color: '57291E' }, line: { type: 'none' } }); // hair

  // audience row — each body is light on the left, shaded on the right
  [{ x: 7.561, body: C.red, dark: 'A71029', hair: '57291E' },
   { x: 8.684, body: C.maroon, dark: '220304', hair: 'B55C46' },
   { x: 9.808, body: C.white, dark: 'BFBFBF', hair: '3B1911' },
   { x: 10.928, body: 'D2D2D2', dark: '9E9E9E', hair: '99780B' }].forEach(a => {
    s.addShape('ellipse', { x: a.x + 0.064, y: 5.036, w: 0.735, h: 0.714, fill: { color: a.hair }, line: { type: 'none' } });
    s.addShape('rect', { x: a.x + 0.352, y: 5.639, w: 0.182, h: 0.193, fill: { color: 'D1A886' }, line: { type: 'none' } });
    s.addShape('round2SameRect', { x: a.x, y: 5.744, w: 0.902, h: 0.906, rectRadius: 0.40, fill: { color: a.dark }, line: { type: 'none' } });
    s.addShape('round1Rect', { x: a.x, y: 5.744, w: 0.451, h: 0.906, rectRadius: 0.40, flipH: true, fill: { color: a.body }, line: { type: 'none' } });
    s.addShape('ellipse', { x: a.x + 0.309, y: 5.70, w: 0.283, h: 0.180, fill: { color: C.white }, line: { type: 'none' } }); // collar
  });
}

// 18 — Visual summary of themes
function slide18(pptx) {
  const s = newSlide(pptx);
  title(s, 'Visual Summary of Themes and Subthemes', 1.036, 0.977, 5.327, 2.051);
  const steps = [{ x: 8.295, y: 1.038, n: '03' }, { x: 5.963, y: 2.533, n: '02' }, { x: 3.672, y: 4.068, n: '01' }];
  steps.forEach(st => {
    s.addShape('ellipse', { x: st.x, y: st.y, w: 0.920, h: 0.920, fill: { color: C.red }, line: { color: C.pink, width: 6 } });
    text(s, 'Step', { x: st.x + 0.146, y: st.y + 0.126, w: 0.622, h: 0.303, align: 'center', fontSize: 12, bold: true, color: C.white });
    text(s, st.n, { x: st.x + 0.137, y: st.y + 0.269, w: 0.647, h: 0.572, align: 'center', fontFace: FONT_HEAD, fontSize: 28, bold: true, color: C.white, wrap: false });
  });
  const bubbles = [{ ox: 8.644, oy: 1.284, tx: 8.775, ty: 1.811, bx: 8.804, by: 2.155, bw: 1.887 },
                   { ox: 6.312, oy: 2.780, tx: 6.443, ty: 3.307, bx: 6.453, by: 3.651, bw: 1.926 },
                   { ox: 4.020, oy: 4.315, tx: 4.152, ty: 4.841, bx: 4.152, by: 5.185, bw: 1.946 }];
  bubbles.forEach(b => {
    s.addShape('ellipse', { x: b.ox, y: b.oy, w: 2.208, h: 2.208, rotate: 7.5, fill: { color: C.white }, line: { type: 'none' }, shadow: SH_SOFT });
    text(s, 'Title Here', { x: b.tx, y: b.ty, w: 1.946, h: 0.337, align: 'center', fontFace: FONT_HEAD, fontSize: 14, bold: true, color: C.red });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ', {
      x: b.bx, y: b.by, w: b.bw, h: 0.811, align: 'center', fontSize: 12, lineSpacingMultiple: 1.2
    });
  });
  [[6.097, 4.556, 0.442, 0.453], [8.389, 2.994, 0.445, 0.323]].forEach(([x, y, w, h]) =>
    s.addShape('line', { x, y, w, h, flipV: true, line: { color: C.rule, width: 1.25, dashType: 'dash' } }));
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis', 8.295, 5.010, 4.002, 0.929);
}

// 19 — Post webinar guide
function slide19(pptx) {
  const s = newSlide(pptx);
  const cols = [
    { ghost: '1', gx: 6.989, gw: 0.628, gc: C.pink, ring: C.red, cx: 7.342, card: 6.587, cardY: 3.581 },
    { ghost: '2', gx: 10.097, gw: 0.837, gc: 'D5CDCD', ring: C.maroon, cx: 10.555, card: 9.800, cardY: 3.607 }
  ];
  cols.forEach(c => {
    text(s, c.ghost, { x: c.gx, y: 1.542, w: c.gw, h: 1.447, align: 'center', fontFace: FONT_HEAD, fontSize: 80, bold: true, color: c.gc });
    s.addShape('ellipse', { x: c.cx, y: 1.621, w: 1.207, h: 1.257, fill: { color: c.ring }, line: { type: 'none' } });
    s.addShape('ellipse', { x: c.cx + 0.156, y: 1.783, w: 0.895, h: 0.932, fill: { color: C.white }, line: { type: 'none' }, shadow: { type: 'outer', blur: 20, offset: 8, angle: 50, color: '000000', opacity: 0.38 } });
    s.addShape('roundRect', { x: c.card, y: c.cardY, w: 2.230, h: 2.871, rectRadius: 0.228, fill: { color: C.white }, line: { type: 'none' }, shadow: SH_SOFT });
  });
  s.addChart(pptx.ChartType.doughnut, [
    { name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [10, 90] }
  ], {
    x: 6.843, y: 3.607, w: 1.719, h: 1.992, chartColors: ['F2F2F2', C.red], holeSize: 75,
    showTitle: false, showLegend: false, showValue: false, dataBorder: { pt: 0, color: C.white }
  });
  text(s, '90%', { x: 7.085, y: 4.231, w: 1.235, h: 0.486, align: 'center', fontFace: FONT_HEAD, fontSize: 18, bold: true, color: '000000', lineSpacingMultiple: 1.4 });
  text(s, 'Successful', { x: 7.151, y: 4.591, w: 1.104, h: 0.315, align: 'center', fontSize: 10, color: C.red, lineSpacingMultiple: 1.4 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer', 6.728, 5.453, 1.948, 0.569, { align: 'center', fontSize: 12 });
  text(s, 'Title Here Two', { x: 9.976, y: 4.151, w: 1.879, h: 0.394, align: 'center', fontFace: FONT_HEAD, fontSize: 14, bold: true, color: C.maroon, lineSpacingMultiple: 1.4 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis', 9.888, 4.545, 2.054, 1.538, { align: 'center', fontSize: 12 });
  iconPie(s, 7.665, 1.969, 0.560, C.red);
  iconVenn(s, 10.878, 1.969, 0.560, C.maroon);
  s.addShape('line', { x: 7.958, y: 3.245, w: 5.375, h: 0, line: { color: C.rule, width: 3, dashType: 'dash' } });
  title(s, 'Post Webinar Guide for Participants', 1.246, 1.309, 4.922, 2.046);
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et', 1.246, 3.710, 4.922, 0.929);
  [{ x: 1.246, y: 4.789, w: 2.406, t: 'Lorem ipsum dolor' },
   { x: 1.246, y: 5.308, w: 2.406, t: 'commodo ligula' },
   { x: 1.246, y: 5.827, w: 2.584, t: 'consectetuer adipiscing' },
   { x: 3.830, y: 4.789, w: 2.584, t: 'natoque penatibus' },
   { x: 3.830, y: 5.308, w: 2.406, t: 'Lorem ipsum dolor' }].forEach(b =>
    para(s, b.t, b.x, b.y, b.w, 0.364, { bullet: { characterCode: '2713', indent: 22 } }));
}

// 20 — Testimonials
function slide20(pptx) {
  const s = newSlide(pptx);
  card(s, 6.788, 4.834, 2.129, 1.188);
  text(s, 'Dennis Hans', { x: 6.951, y: 5.106, w: 1.803, h: 0.404, fontFace: FONT_HEAD, color: C.ink });
  text(s, 'Your Position', { x: 6.951, y: 5.464, w: 1.603, h: 0.286, fontFace: FONT_HEAD, fontSize: 11, italic: true, color: C.ink });
  text(s, '\u201C', { x: 0.62, y: 0.30, w: 1.90, h: 1.90, fontFace: 'Open Sans', fontSize: 200, bold: true, color: C.quote });
  title(s, 'Testimonials from Previous Webinar Participants', 1.246, 1.917, 6.212, 2.046);
  para(s, '\u201C Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce\u201D', 1.228, 4.793, 5.084, 0.646);
  for (let i = 0; i < 5; i++) {
    s.addShape('star5', {
      x: 1.236 + i * 0.3125, y: 5.615, w: 0.259, h: 0.246,
      fill: i < 4 ? { color: C.red } : { type: 'none' }, line: { color: C.red, width: 1.5 }
    });
  }
}

// 21 — Thank you
function slide21(pptx) {
  const s = newSlide(pptx);
  gradient(s, { x: 0, y: 0, w: W, h: H }, [[0, C.redDark], [1, C.maroon]], 44, 56);
  panel(s, { x: 7.986, y: 3.173, w: 2.673, h: 4.327 });
  panel(s, { x: 10.660, y: 0, w: 2.673, h: 3.173 }, 30);
  s.addShape('line', { x: 7.986, y: 0, w: 0, h: H, line: { color: C.red, width: 0.75, transparency: 40 } });
  text(s, 'Thank you', { x: 0.922, y: 4.327, w: 7.102, h: 1.434, fontFace: FONT_HEAD, fontSize: 88, bold: true, color: C.white, lineSpacingMultiple: 0.9 });
  para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa', 0.922, 5.984, 5.469, 0.646, { color: C.white });
  [['www.website.com', 3.353, 2.877], ['12 Your Street Name', 3.677, 2.275], ['+123 456 7890', 4.002, 2.275]]
    .forEach(([t, y, w]) => text(s, t, {
      x: 8.172, y, w, h: 0.325, fontFace: FONT_HEAD, fontSize: 11, color: C.white, lineSpacingMultiple: 1.3
    }));
}

/* --------------------------------------------------------------------- build */
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
                slide08, slide09, slide10, slide11, slide12, slide13, slide14,
                slide15, slide16, slide17, slide18, slide19, slide20, slide21];

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: W, height: H });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: FONT_HEAD, bodyFontFace: FONT_BODY };
  SLIDES.forEach(fn => fn(pptx));
  return pptx.writeFile({ fileName: path.join(__dirname, '01b6198c-baf8-4af3-9343-a12ca0eddefb_grok_final.pptx') });
}

build().then(f => console.log('wrote', f));
