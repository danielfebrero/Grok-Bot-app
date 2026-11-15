/*
 * "Onlino" webinar presentation template - 30 slides, 13.333 x 7.5 in.
 * Rebuilt from scratch with pptxgenjs. Photographic placeholders in the
 * original deck are empty picture frames (they render as nothing) and the
 * four device mock-up photos are redrawn as labelled placeholder blocks.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const C = {
  accent1: '33CCFF', // light cyan     - highlight words, rules, dots
  accent2: '0DC0FF', // mid cyan       - progress bars, pricing header
  accent3: '00AAE6', // blue           - top-left tick, prices
  accent4: '008BBC', // deep blue      - dark end of every gradient
  white: 'FFFFFF',
  dark: '3F3F3F',    // tx1  - headline text
  slate: '595959',   // tx2  - sub headings
  gray: '7F7F7F',    // accent6 - body copy
  grayLt: 'B2B2B2',  // accent6 60% - "The Professional Webinar 2021"
  footer: 'B8B8B8',  // accent6 @55% over white
  footerOnDark: 'D9D9D9',
  cardBg: 'F2F2F2',  // bg2 @35% - pricing card body
  track: 'B2B2B2'    // tx1 40% @35% - progress-bar track
};

const HEAD = 'Cabin Bold';   // theme major font
const BODY = 'Roboto';       // theme minor font

/* ---------------------------------------------------------------- helpers */

function mix (from, to, t) {
  const ch = (hex, i) => parseInt(hex.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map(i => Math.round(ch(from, i) + (ch(to, i) - ch(from, i)) * t))
    .map(v => v.toString(16).padStart(2, '0').toUpperCase())
    .join('');
}

// Clip a polygon (array of [u,v]) against the half plane fn(p) >= 0.
function clipHalf (poly, fn) {
  const out = [];
  poly.forEach((a, i) => {
    const b = poly[(i + 1) % poly.length];
    const fa = fn(a);
    const fb = fn(b);
    if (fa >= 0) out.push(a);
    if ((fa >= 0) !== (fb >= 0)) {
      const t = fa / (fa - fb);
      out.push([a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t]);
    }
  });
  return out;
}

/*
 * Every decorative block in the deck carries the same 315-degree linear
 * gradient: accent1 at the bottom-left corner running to accent4 at the
 * top-right, at a true 45 degrees regardless of the block's proportions.
 * pptxgenjs only knows solid fills, so the block is painted as a stack of
 * diagonal half-planes - the lightest covers everything, each darker one
 * covers a little less of the top-right corner.
 */
function gradient (slide, x, y, w, h, opts) {
  const o = opts || {};
  const N = o.bands || Math.max(12, Math.min(40, Math.round((w + h) * 4)));
  // Position along the ramp, 0 at the light corner, 1 at the dark one.
  const ramp = p => ((o.flipH ? 1 - p[0] : p[0]) * w + (1 - p[1]) * h) / (w + h);
  // `clip` (inches, absolute) keeps only part of the block but not of the ramp.
  const c = o.clip;
  for (let k = 0; k < N; k++) {
    let poly = clipHalf([[0, 0], [1, 0], [1, 1], [0, 1]], p => ramp(p) - k / (N - 1));
    if (c) {
      poly = clipHalf(poly, p => p[0] - (c.x - x) / w);
      poly = clipHalf(poly, p => (c.x + c.w - x) / w - p[0]);
      poly = clipHalf(poly, p => p[1] - (c.y - y) / h);
      poly = clipHalf(poly, p => (c.y + c.h - y) / h - p[1]);
    }
    if (poly.length < 3) continue;
    slide.addShape('custGeom', {
      x: x, y: y, w: w, h: h,
      points: poly
        .map((p, i) => ({ x: p[0] * w, y: p[1] * h, moveTo: i === 0 }))
        .concat([{ close: true }]),
      fill: { color: mix(C.accent1, C.accent4, k / (N - 1)) },
      line: { type: 'none' }
    });
  }
}

function rect (slide, x, y, w, h, color, opts) {
  slide.addShape('rect', Object.assign({
    x: x, y: y, w: w, h: h, fill: { color: color }, line: { type: 'none' }
  }, opts || {}));
}

/*
 * Stand-ins for the four device photographs: a dark bezel around a blank
 * screen, drawn from native shapes instead of embedding the original bitmap.
 */
const SILVER = 'C9CCD0';

/*
 * Opaque device: dark bezel, blank white display, caption.
 * `radius` rounds the outer shell (phones), default is a hard corner.
 */
function deviceScreen (slide, x, y, w, h, bezel, radius) {
  slide.addShape('roundRect', {
    x: x, y: y, w: w, h: h, rectRadius: radius || 0.02,
    fill: { color: '1A1A1A' }, line: { type: 'none' }
  });
  slide.addShape('roundRect', {
    x: x + bezel, y: y + bezel, w: w - 2 * bezel, h: h - 2 * bezel,
    rectRadius: Math.max(0.02, (radius || 0.02) - bezel),
    fill: { color: C.white }, line: { type: 'none' }
  });
  slide.addText('[image]', {
    x: x, y: y + h / 2 - 0.2, w: w, h: 0.4,
    fontFace: BODY, fontSize: 12, color: C.grayLt, align: 'center', valign: 'middle'
  });
}

/*
 * Hollow bezel for the tablet and the laptop: in the original artwork their
 * displays are transparent, so the gradient block behind them shows through.
 * Four strips leave the middle untouched.
 */
function deviceFrame (slide, x, y, w, h, b, color) {
  [[x, y, w, b], [x, y + h - b, w, b],
    [x, y + b, b, h - 2 * b], [x + w - b, y + b, b, h - 2 * b]]
    .forEach(r => {
      slide.addShape('rect', {
        x: r[0], y: r[1], w: r[2], h: r[3],
        fill: { color: color || '1A1A1A' }, line: { type: 'none' }
      });
    });
}

function metalBar (slide, x, y, w, h) {
  slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: SILVER }, line: { type: 'none' } });
}

// Body copy: 11 pt Roboto, justified, 150 % leading, 18 pt space around.
function body (slide, text, x, y, w, h, color, align) {
  slide.addText(text, {
    x: x, y: y, w: w, h: h,
    fontFace: BODY, fontSize: 11, color: color || C.gray,
    align: align || 'justify', valign: 'top',
    lineSpacingMultiple: 1.5, paraSpaceBefore: 18, paraSpaceAfter: 18
  });
}

// 12 pt paragraph used by the "reason / detail" blocks (no space before).
function para12 (slide, text, x, y, w, h, color, opts) {
  slide.addText(text, Object.assign({
    x: x, y: y, w: w, h: h,
    fontFace: BODY, fontSize: 12, color: color || C.gray,
    align: 'justify', valign: 'top', lineSpacingMultiple: 1.5
  }, opts || {}));
}

// 16 pt bold sub heading ("The History Here", "Member Name", ...).
function subHead (slide, text, x, y, w, color, align) {
  slide.addText(text, {
    x: x, y: y, w: w, h: 0.46,
    fontFace: BODY, fontSize: 16, bold: true, color: color || C.slate,
    align: align || 'justify', valign: 'top',
    lineSpacingMultiple: 1.5, paraSpaceBefore: 18, paraSpaceAfter: 18
  });
}

/*
 * Left aligned section header: two-tone 32 pt headline, short accent rule and
 * the "The Professional Webinar 2021" kicker underneath.
 */
function sectionHead (slide, x, y, w, accentText, restText, opts) {
  const o = opts || {};
  slide.addText([
    { text: accentText, options: { color: o.accent || C.accent1 } },
    { text: restText, options: { color: o.rest || C.dark } }
  ], {
    x: x, y: y, w: w, h: 1.178,
    fontFace: HEAD, fontSize: 32, valign: 'top'
  });
  rect(slide, x + 0.122, y + 1.298, 0.5, 0.057, o.rule || C.accent1);
  slide.addText('The Professional Webinar 2021', {
    x: x + 0.657, y: y + 1.125, w: 2.6, h: 0.348,
    fontFace: BODY, fontSize: 11, color: o.kicker || C.grayLt,
    align: 'justify', valign: 'top',
    lineSpacingMultiple: 1.5, paraSpaceBefore: 18, paraSpaceAfter: 18
  });
}

// Centred variant used by the full width slides (pricing, roadmap, ...).
function centerHead (slide, accentText, restText) {
  slide.addText([
    { text: accentText, options: { color: C.accent1 } },
    { text: restText, options: { color: C.dark } }
  ], {
    x: 4.105, y: 1.031, w: 5.123, h: 0.64,
    fontFace: HEAD, fontSize: 32, align: 'center', valign: 'top'
  });
  rect(slide, 5.099, 1.777, 0.5, 0.057, C.accent1);
  slide.addText('The Professional Webinar 2021', {
    x: 5.634, y: 1.603, w: 2.6, h: 0.348,
    fontFace: BODY, fontSize: 11, color: C.grayLt,
    align: 'justify', valign: 'top',
    lineSpacingMultiple: 1.5, paraSpaceBefore: 18, paraSpaceAfter: 18
  });
}

/*
 * Page furniture. `bar`   - accent tick top left
 *                 `sq`    - three little squares top right
 *                 `dots`  - column of four dots on the left edge
 *                 `foot`  - bottom right wordmark (colour depends on backdrop)
 */
function chrome (slide, parts) {
  if (parts.bar) rect(slide, 0.669, 0.737, 0.5, 0.053, C.accent3);
  if (parts.sq) {
    rect(slide, 12.346, 0.762, 0.167, 0.167, C.accent1);
    rect(slide, 12.179, 0.762, 0.167, 0.167, C.accent3);
    rect(slide, 12.512, 0.929, 0.167, 0.167, C.accent3);
  }
  if (parts.dots) {
    [5.379, 5.809, 6.238, 6.668].forEach(y => {
      slide.addShape('ellipse', {
        x: 0.684, y: y, w: 0.096, h: 0.096,
        fill: { color: '18AADB' }, line: { type: 'none' }
      });
    });
  }
  if (parts.foot) {
    slide.addText('Webinar Presentation Template', {
      x: 10.559, y: 6.544, w: 2.195, h: 0.269,
      fontFace: BODY, fontSize: 10, bold: true,
      color: parts.foot === 'light' ? C.footerOnDark : C.footer,
      align: 'right', valign: 'top', wrap: false
    });
  }
}

// Circle bullet + justified paragraph, used on slides 2, 16 and 20.
function bulletRows (slide, textX, dotX, rows) {
  rows.forEach(y => {
    slide.addShape('ellipse', {
      x: dotX, y: y + 0.12, w: 0.386, h: 0.386,
      fill: { color: C.accent1 }, line: { type: 'none' }
    });
    body(slide, TXT.lorem, textX, y, 4.201, 0.903);
  });
}

// "Type Your Skill Here" + track + filled bar + percentage.
function skillBar (slide, x, trackY, pct, label, pctX) {
  slide.addText('Type Your Skill Here', {
    x: x + 0.004, y: trackY - 0.354, w: 2.103, h: 0.269,
    fontFace: BODY, fontSize: 12, italic: true, color: C.gray,
    valign: 'top', margin: 0, lineSpacingMultiple: 1.5
  });
  rect(slide, x, trackY, 4.0, 0.1, C.track, { fill: { color: C.track, transparency: 65 } });
  rect(slide, x, trackY, 4.0 * pct, 0.1, C.accent2);
  slide.addText(label, {
    x: pctX, y: trackY - 0.131, w: 0.5, h: 0.269,
    fontFace: BODY, fontSize: 12, color: C.gray,
    align: 'justify', valign: 'top', margin: 0, lineSpacingMultiple: 1.5
  });
}

// Four-cell grid of "heading + paragraph" blocks.
function infoGrid (slide, cols, rows, heading, text, headW, textW) {
  rows.forEach(y => {
    cols.forEach(x => {
      subHead(slide, heading, x, y, headW || 1.967);
      body(slide, text, x, y + 0.372, textW || 2.377, 0.903);
    });
  });
}

/*
 * Mountain silhouettes of the slide-24 infographic, described as a normalised
 * outline (rounded apex, straight flanks) and scaled into the shape box.
 */
function peak (slide, x, y, w, h, color, apex, rightY) {
  const p = (u, v) => ({ x: u * w, y: v * h });
  const cubic = (u, v, x1, y1, x2, y2) =>
    ({ x: u * w, y: v * h, curve: { type: 'cubic', x1: x1 * w, y1: y1 * h, x2: x2 * w, y2: y2 * h } });
  const l = apex - 0.096;   // left shoulder of the rounded top
  const r = apex + 0.096;
  slide.addShape('custGeom', {
    x: x, y: y, w: w, h: h,
    points: [
      Object.assign(p(apex, 0), { moveTo: true }),
      cubic(r, 0.089, apex + 0.035, 0, apex + 0.070, 0.030),
      p(0.993, rightY), p(1, 1), p(0, 1), p(0.001, 0.994),
      cubic(0.019, 0.947, 0.006, 0.978, 0.012, 0.962),
      p(l, 0.089),
      cubic(apex, 0, l + 0.027, 0.030, l + 0.062, 0),
      { close: true }
    ],
    fill: { color: color }, line: { type: 'none' }
  });
}

/* --------------------------------------------------------------- copy deck */

const TXT = {
  lorem: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Lacus sed viverra tellus in hac',
  loremLong: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Lacus sed viverra tellus in hac habitasse platea dictumst. Rutrum tellus pellentesque eu tincidunt tortor aliquam. Nisl suscipit adipiscing bibendum est tempor incididunt ut labore et',
  loremMed: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Lacus sed viverra tellus in hac habitasse platea dictumst.',
  loremShort: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Lacus sed viverra tellus in hac hitass tempor incid',
  loremTrim: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Lacus sed viverra tellus in ha',
  desc: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqi',
  process: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Lacus sed viverra tellus in hac elit, sed do eiusmod tempor ',
  roadmap: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eimod tempor incididunt ut labore et dolore magna aliqua. Lacu',
  adip: 'Adipiscing elit pellentesque habitant morbi tristique sen ectus et netus. Int eg er',
  adipTopic: 'Adipiscing elit pellentes que habitant morbi tristique sen ectus et netus. Int eg er',
  adipPhil: 'Adipiscing elit pellentesque hab itant morbi tristique sen ectus et netus. Int eg er platea ',
  skill: 'Adipiscing elit pellentesque habitant morbi tristique senectus et netus. Integer enim neque volutpat ac tincidunt vitae semper. Sit amet mauris commodo quis. A pellentesque sit amet',
  member: 'Adipiscing elit pellentesque habitant morbi ttique ',
  proin1: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin sed libero in magna ultrices gravida sit ametat diam. ',
  proin2: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin sed libero in magna ultrices gravida sit ametat diam. Suspen disse placerat gravida magna vel fermentum iosla',
  infoWide: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Lacus sed viverra tellus in hac habitasse platea dictumst. Rutrum tellus pellentesque eu tincidunt tortor aliquam. Nisl suscipit adipiscing bibendum est tempor incididunt ut labore et Lacus sed viverra tellus in hac habitasse platea ',
  priceLine: 'Lorem ipsum dolor sit amet, cons'
};

/* --------------------------------------------------------- slide builders */

// 1 - cover
function slide01 (s) {
  gradient(s, 9.649, 3.75, 3.684, 3.75);
  chrome(s, { bar: 1, sq: 1, dots: 1, foot: 'light' });
  s.addText('Webinar Presentation Template', {
    x: 1.584, y: 1.902, w: 4.201, h: 0.415,
    fontFace: BODY, fontSize: 14, bold: true, color: C.gray,
    align: 'justify', valign: 'top', lineSpacingMultiple: 1.5, paraSpaceBefore: 18
  });
  s.addText([
    { text: 'On', options: { color: C.accent1 } },
    { text: 'lino', options: { color: C.dark } }
  ], { x: 1.509, y: 2.052, w: 5.193, h: 2.036, fontFace: HEAD, fontSize: 115, valign: 'top' });
  rect(s, 3.053, 3.961, 5.193, 0.162, C.accent1);
  body(s, TXT.lorem, 1.544, 4.632, 4.201, 0.903);
}

// 2 - about our missions
function slide02 (s) {
  gradient(s, 2.625, 3.264, 4.042, 4.236);
  chrome(s, { sq: 1, foot: 1 });
  sectionHead(s, 7.041, 1.28, 3.518, 'About Our ', 'Missions Here.');
  bulletRows(s, 7.616, 7.171, [3.212, 4.28, 5.347]);
}

// 3 - general description
function slide03 (s) {
  gradient(s, 9.688, 0, 3.062, 4.792, 0, true);
  chrome(s, { bar: 1, dots: 1 });
  sectionHead(s, 1.572, 1.613, 3.867, 'About General ', 'Description Here.');
  body(s, TXT.loremLong, 1.59, 3.755, 4.78, 1.459);
  body(s, TXT.loremShort, 1.59, 5.299, 4.78, 0.903);
}

// 4 - our history (2 x 2 grid)
function slide04 (s) {
  gradient(s, 0, 2.75, 3.708, 4.0);
  chrome(s, { sq: 1, foot: 1 });
  sectionHead(s, 6.574, 1.227, 3.518, 'About Our ', 'History Here.');
  infoGrid(s, [6.574, 9.325], [3.151, 4.613], 'The History Here', TXT.adip);
}

// 5 - our philosophy (white text over the bottom band)
function slide05 (s) {
  gradient(s, 0, 4.75, 13.333, 2.75);
  chrome(s, { bar: 1, sq: 1 });
  sectionHead(s, 1.572, 1.613, 3.867, 'About Our ', 'Philosophy Here.');
  body(s, TXT.loremMed, 1.572, 3.525, 4.95, 0.903);
  [1.574, 4.219].forEach(x => {
    subHead(s, 'The Philosophy', x, 5.047, 1.967, C.white);
    body(s, TXT.adipPhil, x, 5.418, 2.377, 0.903, C.white);
  });
}

// 6 - we always bring the solutions
function slide06 (s) {
  gradient(s, 2.604, 0, 3.062, 4.812);
  chrome(s, { sq: 1, foot: 1 });
  sectionHead(s, 6.592, 1.525, 3.518, 'We Always Bring ', 'The Solutions.');
  para12(s, TXT.proin1, 6.574, 3.497, 4.756, 0.673);
  para12(s, TXT.proin2, 6.574, 4.245, 4.756, 0.976);
  para12(s, 'THE REASON HERE', 6.574, 5.457, 2.6, 0.415, C.slate, { fontSize: 14, bold: true });
  para12(s, 'The Detail Of Reason Right Here', 6.574, 5.748, 2.761, 0.37);
}

// 7 - our services (white copy over the bottom band)
function slide07 (s) {
  gradient(s, 0, 3.596, 13.333, 3.904);
  chrome(s, { bar: 1, dots: 1 });
  sectionHead(s, 1.572, 1.613, 3.867, 'About Our ', 'Services Here.');
  [['THE REASON ONE HERE', 3.866, 3.563],
    ['THE REASON TWO HERE', 4.228, 3.113],
    ['THE REASON THREE HERE', 4.523, 3.313]].forEach(r => {
    para12(s, r[0], 1.624, r[1], r[2], 0.37, C.white,
      { bold: true, bullet: { indent: 22.5 } });
  });
  para12(s, TXT.proin1, 1.624, 5.069, 4.756, 0.673, C.white);
  para12(s, TXT.proin2, 1.624, 5.804, 4.756, 0.976, C.white);
}

// 8 - the solution for your event (white header over the top band)
function slide08 (s) {
  gradient(s, 0, 0.012, 13.333, 3.75);
  chrome(s, { bar: 1, dots: 1, foot: 1 });
  sectionHead(s, 1.572, 1.613, 3.257, 'The Solution For Your Event.', '',
    { accent: C.white, rule: C.white, kicker: C.white });
  [3.913, 5.21].forEach(y => {
    subHead(s, 'The Description Here', 1.558, y, 3.446);
    body(s, TXT.desc, 1.572, y + 0.403, 4.353, 0.625);
  });
}

// 9 - we bring the future era
function slide09 (s) {
  gradient(s, 8.646, 0, 4.021, 4.812);
  chrome(s, { bar: 1, dots: 1 });
  sectionHead(s, 1.539, 1.541, 3.518, 'We Bring The ', 'Future Era.');
  infoGrid(s, [1.539, 4.114], [3.499, 4.927], 'The Reason Here', TXT.adip);
}

// 10 - inner member (two gradient cards + three skill bars)
function slide10 (s) {
  gradient(s, 6.736, 1.158, 2.636, 2.592);
  gradient(s, 9.37, 3.75, 2.636, 2.592);
  chrome(s, { bar: 1, dots: 1 });
  sectionHead(s, 1.539, 1.541, 3.518, 'Meet Our Inner ', 'Member Here.');
  body(s, TXT.skill, 1.575, 3.328, 4.828, 0.903);
  skillBar(s, 1.706, 4.736, 0.75, '75%', 5.967);
  skillBar(s, 1.706, 5.378, 0.875, '90%', 5.967);
  skillBar(s, 1.706, 6.02, 0.75, '75%', 5.967);
  [[7.07, 1.982, 6.865, 2.441], [9.705, 4.574, 9.5, 5.033]].forEach(m => {
    subHead(s, 'Member Name', m[0], m[1], 1.967, C.white, 'center');
    body(s, TXT.member, m[2], m[3], 2.377, 0.625, C.white, 'center');
  });
}

// 11 - senior planners
function slide11 (s) {
  gradient(s, 0.688, 0, 3.0, 3.75);
  chrome(s, { sq: 1, foot: 1 });
  sectionHead(s, 6.541, 1.541, 3.518, 'Meet Our Senior ', 'Planner Here.');
  infoGrid(s, [6.558, 9.23], [3.499, 4.927], 'Member Name', TXT.adip);
}

// 12 - general manager
function slide12 (s) {
  gradient(s, 8.604, 3.75, 4.729, 3.0);
  chrome(s, { bar: 1, dots: 1 });
  sectionHead(s, 1.539, 1.541, 3.829, 'Meet Onlino\u2019s  ', 'General Manager.');
  subHead(s, 'Member Name', 1.556, 3.529, 1.967);
  body(s, TXT.skill, 1.575, 3.936, 4.828, 0.903);
  skillBar(s, 1.688, 5.379, 0.75, '75%', 5.949);
  skillBar(s, 1.688, 6.021, 0.875, '90%', 5.949);
}

// 13 - ceo and founder
function slide13 (s) {
  gradient(s, 0, 3.627, 6.667, 3.021);
  chrome(s, { sq: 1, foot: 1 });
  sectionHead(s, 6.879, 1.541, 3.829, 'Meet Onlino\u2019s  ', 'CEO And Founder.');
  subHead(s, 'Member Name', 6.897, 3.564, 1.967);
  body(s, TXT.skill, 6.916, 3.971, 4.828, 0.903);
  skillBar(s, 7.029, 5.414, 0.75, '75%', 11.289);
  skillBar(s, 7.029, 6.056, 0.875, '90%', 11.289);
}

// 14 - section break
function slide14 (s) {
  gradient(s, 8.596, 2.825, 4.737, 3.037);
  chrome(s, { bar: 1, dots: 1 });
  rect(s, 1.684, 2.07, 1.5, 0.1, C.accent1);
  s.addText([
    { text: 'It\u2019s Time ', options: { color: C.dark } },
    { text: 'To Break', options: { color: C.accent1 } }
  ], { x: 1.536, y: 2.127, w: 4.464, h: 2.322, fontFace: HEAD, fontSize: 66, valign: 'top' });
  body(s, TXT.loremMed, 1.596, 4.959, 4.78, 0.903);
}

// 15 - best gallery
function slide15 (s) {
  gradient(s, 6.674, 0.789, 6.659, 2.961);
  chrome(s, { bar: 1, dots: 1 });
  sectionHead(s, 1.572, 1.613, 3.409, 'About Our Best ', 'Gallery Here.');
  body(s, TXT.loremLong, 1.59, 3.702, 4.78, 1.459);
  body(s, TXT.loremShort, 1.59, 5.246, 4.78, 0.903);
}

// 16 - portfolio
function slide16 (s) {
  gradient(s, 3.729, 3.75, 3.0, 3.021);
  chrome(s, { sq: 1, foot: 1 });
  sectionHead(s, 7.041, 1.367, 3.518, 'About Our ', 'Portfolio Here.');
  bulletRows(s, 7.616, 7.171, [3.212, 4.28, 5.347]);
}

// 17 - process portfolio (centred header, white copy on the block)
function slide17 (s) {
  gradient(s, 0, 2.698, 6.667, 4.093);
  chrome(s, { bar: 1, sq: 1 });
  centerHead(s, 'The Process ', 'Portfolio.');
  [3.274, 4.91].forEach(y => {
    subHead(s, 'The Description Here', 0.669, y, 3.446, C.white);
    body(s, TXT.process, 0.684, y + 0.403, 5.018, 0.903, C.white);
  });
}

// 18 - future orientations
function slide18 (s) {
  gradient(s, 0.646, 1.729, 3.083, 5.771);
  chrome(s, { sq: 1, foot: 1 });
  sectionHead(s, 6.967, 1.525, 4.19, 'Let\u2019s Start The ', 'Future Orientations.');
  para12(s, TXT.proin1, 6.95, 3.532, 4.756, 0.673);
  para12(s, TXT.proin2, 6.95, 4.28, 4.756, 0.976);
  para12(s, 'THE REASON HERE', 6.95, 5.492, 2.6, 0.415, C.slate, { fontSize: 14, bold: true });
  para12(s, 'The Detail Of Reason Right Here', 6.95, 5.783, 2.761, 0.37);
}

// 19 - pricing tables
function slide19 (s) {
  chrome(s, { bar: 1, sq: 1, dots: 1, foot: 1 });
  centerHead(s, 'Pricing ', 'Tables.');
  [['Basic Plan', '2', 2.412], ['Standart Plan', '5', 5.336], ['Premium Plan', '7', 8.26]]
    .forEach(card => {
      const name = card[0];
      const lead = card[1];
      const x = card[2];
      rect(s, x, 2.314, 2.661, 3.961, C.cardBg, { fill: { color: C.cardBg, transparency: 65 } });
      rect(s, x, 2.314, 2.661, 0.824, C.accent2);
      s.addText(name, {
        x: x + 0.368, y: 2.361, w: 1.924, h: 0.505,
        fontFace: BODY, fontSize: 18, bold: true, color: C.white,
        align: 'center', valign: 'top', lineSpacingMultiple: 1.5
      });
      s.addText('Monthly Plan', {
        x: x + 0.754, y: 2.701, w: 1.152, h: 0.348,
        fontFace: BODY, fontSize: 11, color: C.white,
        align: 'center', valign: 'top', lineSpacingMultiple: 1.5
      });
      s.addText('Start From', {
        x: x + 0.665, y: 3.298, w: 1.331, h: 0.46,
        fontFace: BODY, fontSize: 16, bold: true, color: C.gray,
        align: 'center', valign: 'top', lineSpacingMultiple: 1.5
      });
      s.addText([
        { text: '$', options: { fontSize: 28 } },
        { text: lead + '99', options: { fontSize: 36 } }
      ], {
        x: x + 0.72, y: 3.712, w: 1.22, h: 0.707,
        fontFace: BODY, bold: true, color: C.accent3,
        align: 'center', valign: 'top', wrap: false
      });
      [4.538, 4.879, 5.221].forEach(y => {
        s.addText(TXT.priceLine, {
          x: x + 0.14, y: y, w: 2.38, h: 0.337,
          fontFace: BODY, fontSize: 10, color: C.gray,
          align: 'center', valign: 'top', lineSpacingMultiple: 1.5
        });
      });
      s.addShape('rect', {
        x: x + 0.841, y: 5.747, w: 0.979, h: 0.245,
        fill: { color: C.accent2 }, line: { type: 'none' }
      });
      s.addText('Get It Now', {
        x: x + 0.841, y: 5.747, w: 0.979, h: 0.245,
        fontFace: BODY, fontSize: 10, color: C.white, align: 'center', valign: 'middle'
      });
    });
}

// 20 - monitor mock-up
function slide20 (s) {
  chrome(s, { bar: 1, sq: 1, dots: 1, foot: 1 });
  deviceScreen(s, 1.583, 2.153, 4.528, 2.694, 0.18, 0.06);
  metalBar(s, 1.583, 4.847, 4.528, 0.375);   // brushed chin
  metalBar(s, 3.389, 5.222, 0.917, 0.611);   // stand neck
  metalBar(s, 3.167, 5.833, 1.361, 0.09);    // stand foot
  sectionHead(s, 6.988, 1.28, 3.518, 'About Monitor ', 'Mockup Devices.');
  bulletRows(s, 7.563, 7.118, [3.212, 4.28, 5.347]);
}

// 21 - tablet mock-up
function slide21 (s) {
  gradient(s, 8.702, 0.789, 4.632, 2.961);
  deviceFrame(s, 7.694, 1.736, 3.458, 4.778, 0.155);
  s.addShape('roundRect', {                  // stylus clipped to the edge
    x: 11.056, y: 1.903, w: 0.181, h: 4.472, rectRadius: 0.09,
    fill: { color: 'F2F2F2' }, line: { color: SILVER, width: 1 }
  });
  chrome(s, { bar: 1, dots: 1 });
  sectionHead(s, 1.539, 1.541, 3.518, 'About Tablet ', 'Mockup Devices.');
  infoGrid(s, [1.539, 4.114], [3.499, 4.927], 'The Topic Here', TXT.adipTopic);
}

// Rear of the phone pair on slide 22 - silver body with the camera bump.
function slide22Body (s, x, y, w, h) {
  s.addShape('roundRect', {
    x: x, y: y, w: w, h: h, rectRadius: 0.36,
    fill: { color: 'E9E9E7' }, line: { color: SILVER, width: 1 }
  });
  s.addShape('roundRect', {
    x: x + 0.11, y: y + 0.13, w: 0.29, h: 0.62, rectRadius: 0.09,
    fill: { color: 'D5D5D3' }, line: { type: 'none' }
  });
  [0.2, 0.44].forEach(dy => {
    s.addShape('ellipse', {
      x: x + 0.17, y: y + dy, w: 0.17, h: 0.17,
      fill: { color: '3A3A3A' }, line: { type: 'none' }
    });
  });
}

// 22 - phone mock-up
function slide22 (s) {
  gradient(s, 1.229, 3.75, 2.999, 3.75);
  slide22Body(s, 2.333, 1.153, 1.75, 5.181);  // rear phone, silver body
  deviceScreen(s, 3.125, 1.167, 2.528, 5.194, 0.075, 0.35);
  chrome(s, { sq: 1, foot: 1 });
  sectionHead(s, 6.592, 1.525, 3.518, 'About Phone ', 'Mockup Devices.');
  para12(s, TXT.proin1, 6.574, 3.497, 4.756, 0.673);
  para12(s, TXT.proin2, 6.574, 4.245, 4.756, 0.976);
  para12(s, 'THE REASON HERE', 6.574, 5.457, 2.6, 0.415, C.slate, { fontSize: 14, bold: true });
  para12(s, 'The Detail Of Reason Right Here', 6.574, 5.748, 2.761, 0.37);
}

// 23 - laptop mock-up
function slide23 (s) {
  gradient(s, 8.667, 0, 4.667, 7.5);
  deviceFrame(s, 7.222, 2.833, 4.722, 3.111, 0.042, SILVER); // lid shell rim
  deviceFrame(s, 7.264, 2.875, 4.639, 3.028, 0.125);
  metalBar(s, 6.708, 5.972, 5.681, 0.181);   // laptop base
  chrome(s, { bar: 1, dots: 1 });
  sectionHead(s, 1.539, 1.541, 3.518, 'About Laptop ', 'Mockup Devices.');
  infoGrid(s, [1.539, 4.114], [3.499, 4.927], 'The Topic Here', TXT.adipTopic);
}

// 24 - mountain infographic
function slide24 (s) {
  chrome(s, { bar: 1, sq: 1, dots: 1 });
  peak(s, 5.222, 2.871, 5.611, 4.629, C.accent3, 0.498, 0.982);
  peak(s, 7.784, 3.54, 4.8, 3.96, C.accent4, 0.498, 0.982);
  peak(s, 9.94, 4.241, 3.394, 3.259, C.accent3, 0.58, 0.682);
  [[6.767, '82%'], [8.849, '65%'], [11.021, '50%']].forEach(d => {
    s.addText('Data Here', {
      x: d[0], y: 6.073, w: 1.758, h: 0.444,
      fontFace: BODY, fontSize: 18, bold: true, color: C.white,
      align: 'center', valign: 'top', lineSpacingMultiple: 1.2
    });
  });
  [['82%', 7.244, 1.185], ['65%', 9.397, 1.84], ['50%', 11.132, 2.452]].forEach(v => {
    s.addText(v[0], {
      x: v[1], y: v[2], w: 1.566, h: 0.404,
      fontFace: HEAD, fontSize: 18, color: C.gray, align: 'center', valign: 'top'
    });
    s.addText('A wonderful serenity has taken.', {
      x: v[1], y: v[2] + 0.33, w: 1.566, h: 0.53,
      fontFace: BODY, fontSize: 11, color: C.gray,
      align: 'center', valign: 'top', lineSpacingMultiple: 1.2
    });
  });
  [[7.906, 2.11, C.accent3], [10.069, 2.813, C.accent4], [11.789, 3.465, C.accent3]]
    .forEach(o => {
      s.addShape('ellipse', {
        x: o[0], y: o[1], w: 0.222, h: 0.222,
        fill: { color: o[2] }, line: { type: 'none' }
      });
    });
  sectionHead(s, 1.572, 1.578, 3.409, 'About Company ', 'Infographic.');
  body(s, TXT.loremLong, 1.59, 3.597, 4.78, 1.459);
  body(s, TXT.loremTrim, 1.59, 5.141, 4.126, 0.903);
}

// 25 - teardrop infographic
function slide25 (s) {
  chrome(s, { bar: 1, sq: 1, dots: 1, foot: 1 });
  centerHead(s, 'About ', 'Infographic.');
  [[2.716, C.accent1], [4.776, '00ACE5'], [6.817, C.accent3], [8.824, '007FAC']]
    .forEach((t, i) => {
      s.addShape('teardrop', {
        x: t[0], y: 3.169, w: 1.793, h: 1.793, rotate: 45,
        fill: { color: t[1] }, line: { type: 'none' },
        shadow: { type: 'outer', color: '000000', opacity: 0.26, blur: 31, offset: 3, angle: 55 }
      });
      // white pictogram sitting in the round part of the drop
      s.addShape('roundRect', {
        x: t[0] + 0.67, y: 3.755, w: 0.46, h: 0.62, rectRadius: 0.12,
        fill: { color: C.white }, line: { type: 'none' }
      });
      s.addText('Topic Here', {
        x: [2.827, 4.886, 6.928, 8.934][i], y: 2.516, w: 1.572, h: 0.46,
        fontFace: BODY, fontSize: 16, bold: true, color: C.slate,
        align: 'center', valign: 'top', lineSpacingMultiple: 1.5
      });
    });
  body(s, TXT.infoWide, 1.739, 5.226, 9.856, 0.903, C.gray, 'center');
}

// 26 / 27 / 28 - roadmap pages, only the bar width and the dates change
function roadmap (s, barX, barW, dates) {
  gradient(s, barX, 3.702, barW, 0.095, 30);
  chrome(s, { bar: 1, sq: 1, dots: 1, foot: 1 });
  centerHead(s, 'Company ', 'Roadmap.');
  [1.588, 7.062].forEach((x, i) => {
    subHead(s, dates[i], x, 4.938, 3.254);
    body(s, TXT.roadmap, x, 5.31, 4.522, 0.625);
  });
}

// 29 - contact information
function slide29 (s) {
  rect(s, 8.692, 4.678, 4.641, 2.085, C.accent1);
  chrome(s, { bar: 1, sq: 1, dots: 1 });
  sectionHead(s, 1.572, 1.613, 3.409, 'Contact ', 'Informations.');
  [['Office Hours', 1.572, 3.542, 2.6, ['Monday \u2013 Saturday', '08.00 AM \u2013 08.00 PM'], 2.6, 'justify'],
    ['Address', 3.894, 3.542, 1.528, ['Marketing Pop-Up, 4 - 5 Goodwin Street, N4 3HQ'], 2.208, 'left'],
    ['Get In Touch', 1.572, 4.9, 2.6, ['(+62) 000 0000 0000', '(0341) 00000'], 2.6, 'justify'],
    ['Follow Us', 3.894, 4.9, 1.528, ['www.onlino.com', 'office@onlino.com'], 2.6, 'left']]
    .forEach(f => {
      para12(s, f[0], f[1], f[2], f[3], 0.505, C.slate, { fontSize: 18, bold: true });
      para12(s, f[4].join('\n'), f[1], f[2] + 0.463, f[5], 0.673, C.gray, { align: f[6] });
    });
}

// 30 - thanks
function slide30 (s) {
  gradient(s, 9.649, 3.75, 3.684, 3.75);
  chrome(s, { bar: 1, sq: 1, dots: 1, foot: 'light' });
  s.addText('It\u2019s Time To Say Good Bye', {
    x: 1.356, y: 1.902, w: 4.201, h: 0.415,
    fontFace: BODY, fontSize: 14, bold: true, color: C.gray,
    align: 'justify', valign: 'top', lineSpacingMultiple: 1.5, paraSpaceBefore: 18
  });
  s.addText([
    { text: 'Th', options: { color: C.accent1 } },
    { text: 'anks', options: { color: C.dark } }
  ], { x: 1.333, y: 2.052, w: 5.193, h: 2.036, fontFace: HEAD, fontSize: 115, valign: 'top' });
  rect(s, 3.053, 3.961, 5.193, 0.162, C.accent1);
  body(s, TXT.lorem, 1.368, 4.632, 4.201, 0.903);
}

const DECK = [
  slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
  slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25,
  s => roadmap(s, 3.233, 10.1, ['Roadmap, 24 Jan 2020', 'Roadmap, 09 Feb 2020']),
  s => roadmap(s, 0, 13.333, ['Roadmap, 10 Apr 2021', 'Roadmap, 29 Apr 2020']),
  s => roadmap(s, 0, 10.3, ['Roadmap, 19 Jun 2021', 'Roadmap, 07 Jul 2021']),
  slide29, slide30
];

/* -------------------------------------------------------------------- main */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pptx.title = 'Onlino - Webinar Presentation Template';

DECK.forEach(build => {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  build(slide);
});

pptx.writeFile({
  fileName: path.join(__dirname, '038389f1-7513-4243-ac4a-9e6b33024d43_grok_final.pptx')
});
