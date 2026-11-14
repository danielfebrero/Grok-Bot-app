/*
 * Yogathy — yoga studio pitch deck (62 slides, 13.333in x 7.5in widescreen).
 *
 * Rebuilt from scratch with pptxgenjs: every slide is a builder function that
 * calls the small vocabulary of helpers defined at the top of this file.
 * The reference deck's photos live in picture placeholders; here they are
 * replaced by flat "[image]" placeholder rectangles (see `photo`).
 *
 * Run:  node 098c4285-8e5b-46f0-9d7c-eb0efe97f6c2_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const GREEN = '2EB34A', LIME = '51D36D', MINT = '7CDE91', FOREST = '26923D';
const AMBER = 'FFC000', GOLD = 'FFCF37', BRONZE = 'E2AC00', SAND = 'FFDC6D';
const W = 'FFFFFF', BLACK = '000000', INK = '262626', GREY = '808080';
const SILVER = 'D9D9D9', MIST = 'F2F2F2', CLOUD = 'E7E6E6', HAZE = 'EBEBEB';
const ASH = 'BFBFBF', STONE = 'A6A6A6';
const NAVY = '3F3D56', SLATE = '464353', CHAR = '35323E';
const SKIN = 'DB8B8B', ROSE = 'F86D70', VIOLET = '6C63FF', PLUM = '6A5C99';

/* --------------------------------------------------------------- typography */
const WS = 'Work Sans', WSM = 'Work Sans Medium', OS = 'Open Sans';
const PP = 'Poppins', PSB = 'Poppins SemiBold', NS = 'Noto Sans';

/* ------------------------------------------------------------------ effects */
/* pptxgenjs rewrites the shadow object it is handed, so presets are factories. */
const SHADOW      = () => ({ type: 'outer', blur: 19, offset: 0.01, angle: 0, color: BLACK, opacity: 0.17 });
const SHADOW_WIDE = () => ({ type: 'outer', blur: 23, offset: 0.01, angle: 0, color: BLACK, opacity: 0.20 });
const SHADOW_GREY = () => ({ type: 'outer', blur: 21, offset: 0.01, angle: 0, color: GREY,  opacity: 0.42 });
const SHADOW_DROP = () => ({ type: 'outer', blur: 4,  offset: 3,    angle: 45, color: BLACK, opacity: 0.40 });

/** average of two hex colours - stands in for a two-stop gradient fill */
function mid(a, b) {
  const p = (i) => Math.round((parseInt(a.substr(i, 2), 16) + parseInt(b.substr(i, 2), 16)) / 2);
  return [p(0), p(2), p(4)].map((v) => v.toString(16).toUpperCase().padStart(2, '0')).join('');
}

const NO_LINE = { type: 'none' };

/* ------------------------------------------------------------------ helpers */

/** generic shape */
function sh(s, kind, x, y, w, h, o) {
  const opt = Object.assign({ x, y, w, h, line: NO_LINE }, o || {});
  if (typeof opt.shadow === 'function') opt.shadow = opt.shadow();
  s.addShape(kind, opt);
}

/** straight horizontal rule */
function line(s, x, y, w, color, pt) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width: pt } });
}

/**
 * Soft one-directional colour wash: the reference uses an alpha gradient,
 * rebuilt here as a stack of increasingly opaque bands.
 * `start` = fraction of the box where the colour begins, `peak` = final alpha.
 */
function fade(s, x, y, w, h, color, start, peak, dir) {
  const BANDS = 14;
  for (let i = 0; i < BANDS; i++) {
    const t0 = start + ((1 - start) * i) / BANDS;
    const t1 = start + ((1 - start) * (i + 1)) / BANDS;
    const alpha = peak * ((i + 1) / BANDS);
    if (dir === 'h') {
      sh(s, 'rect', x + w * t0, y, w * (t1 - t0) + 0.01, h,
        { fill: { color, transparency: Math.round(100 - alpha * 100) } });
    } else {
      sh(s, 'rect', x, y + h * t0, w, h * (t1 - t0) + 0.01,
        { fill: { color, transparency: Math.round(100 - alpha * 100) } });
    }
  }
}

/**
 * Same wash, but clipped to a four-point quad: each band is a slice of the
 * quad, so sheared panels keep their diagonal edges.
 */
function fadeQuad(s, x, y, w, h, quad, color, start, peak) {
  const BANDS = 14;
  const lerp = (a, b, t) => a + (b - a) * t;
  for (let i = 0; i < BANDS; i++) {
    const t0 = start + ((1 - start) * i) / BANDS;
    const t1 = start + ((1 - start) * (i + 1)) / BANDS;
    const pts = [
      { x: lerp(quad[0].x, quad[3].x, t0), y: h * t0 },
      { x: lerp(quad[1].x, quad[2].x, t0), y: h * t0 },
      { x: lerp(quad[1].x, quad[2].x, t1), y: h * t1 + 0.01 },
      { x: lerp(quad[0].x, quad[3].x, t1), y: h * t1 + 0.01 },
    ];
    sh(s, 'custGeom', x, y, w, h, {
      fill: { color, transparency: Math.round(100 - peak * ((i + 1) / BANDS) * 100) },
      points: pts,
    });
  }
}

/** placeholder standing in for a photo in the reference deck */
function photo(s, x, y, w, h, o) {
  sh(s, 'rect', x, y, w, h, Object.assign({ fill: { color: HAZE } }, o || {}));
  s.addText('[image]', {
    x, y: y + h / 2 - 0.18, w, h: 0.36, align: 'center', valign: 'middle',
    fontFace: WS, fontSize: 11, color: STONE,
  });
}

/* --------------------------------------------------------------- text atoms */

/**
 * Text runs are written compactly as `[text, color, extraOptions]` tuples so
 * that mixed-colour headings stay readable; `runs` expands them for pptxgenjs.
 * A plain string is passed through untouched.
 */
function runs(text, color) {
  if (!Array.isArray(text)) return text;
  return text.map((r) => ({
    text: r[0],
    options: Object.assign({ color: r[1] || color }, r[2] || {}),
  }));
}

/** small bold Poppins eyebrow above a heading */
function kick(s, x, y, w, text, color, align) {
  s.addText(runs(text, color), { x, y, w, h: 0.303, fontFace: PP, fontSize: 12, bold: true,
    charSpacing: 1, color, align: align || 'left', valign: 'middle', wrap: false });
}
function kick14(s, x, y, w, text, color, align) {
  s.addText(runs(text, color), { x, y, w, h: 0.337, fontFace: PP, fontSize: 14, bold: true,
    charSpacing: 1, color, align: align || 'left', valign: 'middle', wrap: false });
}
/** 16pt Poppins tagline used under the cover wordmarks */
function tag(s, x, y, w, text, color, align) {
  s.addText(runs(text, color), { x, y, w, h: 0.37, fontFace: PP, fontSize: 16, bold: true,
    charSpacing: 2, color, align: align || 'left', valign: 'middle', wrap: false });
}
/** Poppins label with inherited size (used inside pills) */
function chip(s, x, y, w, text, color, align) {
  s.addText(runs(text, color), { x, y, w, h: 0.404, fontFace: PP, fontSize: 18, bold: true,
    charSpacing: 1, color, align: align || 'left', valign: 'middle', wrap: false });
}
/** 29pt Open Sans section heading (mixed-colour runs allowed) */
function head(s, x, y, w, h, text, align, color) {
  s.addText(runs(text, color || INK), { x, y, w, h, fontFace: OS, fontSize: 29, bold: true,
    color: color || INK, align: align || 'left', valign: 'top' });
}
/** 88pt Work Sans cover wordmark */
function hero(s, x, y, w, h, text, color, align) {
  s.addText(runs(text, color), { x, y, w, h, fontFace: WS, fontSize: 88, bold: true,
    charSpacing: 6, color, align: align || 'left', valign: 'middle', wrap: false });
}
/** 44pt Open Sans statistic */
function num44(s, x, y, w, h, text, color, align) {
  s.addText(runs(text, color), { x, y, w, h, fontFace: OS, fontSize: 44, bold: true,
    color, align: align || 'left', valign: 'middle' });
}
/** 11pt Work Sans justified body copy on 1.5 leading */
function body(s, x, y, w, h, text, color, align) {
  s.addText(runs(text, color), { x, y, w, h, fontFace: WS, fontSize: 11, color,
    align: align || 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
}
/** 11pt Open Sans supporting copy */
function small(s, x, y, w, h, text, color, align) {
  s.addText(runs(text, color), { x, y, w, h, fontFace: OS, fontSize: 11, color,
    align: align || 'left', lineSpacingMultiple: 1.5, valign: 'top' });
}
/** trainer / member names */
function name(s, x, y, w, h, text, color, align) {
  s.addText(runs(text, color), { x, y, w, h, fontFace: PSB, fontSize: 18, color,
    align: align || 'left', valign: 'middle', wrap: false });
}
/** top-bar navigation words */
function navlink(s, x, y, w, text, color) {
  s.addText(runs(text, color), { x, y, w, h: 0.303, fontFace: WSM, fontSize: 12, color,
    valign: 'middle', margin: 0 });
}
/** escape hatch for one-off text formatting */
function txt(s, x, y, w, h, text, o) {
  s.addText(runs(text, o.color), { x, y, w, h, fontFace: o.face || WS, fontSize: o.size || 11,
    bold: !!o.bold, italic: !!o.italic, charSpacing: o.spc, color: o.color,
    align: o.align || 'left', lineSpacingMultiple: o.lsp,
    valign: o.valign || 'top', wrap: o.wrap !== false });
}

/* ------------------------------------------------------------ brand pieces */

/** heart-with-pulse studio mark */
function logo(s, x, y, color) {
  sh(s, 'heart', x, y, 0.451, 0.428, { fill: { color }, rotate: 180 });
  sh(s, 'custGeom', x + 0.07, y + 0.15, 0.31, 0.16, {
    fill: { type: 'none' }, line: { color: W, width: 1.25 },
    points: [{ x: 0, y: 0.09 }, { x: 0.08, y: 0.09 }, { x: 0.13, y: 0.0 },
             { x: 0.18, y: 0.16 }, { x: 0.23, y: 0.09 }, { x: 0.31, y: 0.09 }],
  });
}

/** site address, always the same size */
function url(s, x, y, color) {
  s.addText('www.yogathy.com', { x, y, w: 1.752, h: 0.349, fontFace: NS,
    fontSize: 11, color, lineSpacingMultiple: 1.5, valign: 'top' });
}

/** "2020 —— Yoga Healthy Corp." corner credit */
function stamp(s, x, y, color) {
  s.addText('2020', { x, y, w: 0.637, h: 0.337, fontFace: NS, fontSize: 14,
    italic: true, color, valign: 'middle' });
  line(s, x + 0.623, y + 0.166, 0.677, color, 1);
  s.addText('Yoga Healthy Corp.', { x: x + 1.421, y, w: 1.874, h: 0.337,
    fontFace: NS, fontSize: 14, italic: true, color, valign: 'middle' });
}

/** facebook / youtube / linkedin circles */
function social(s, x, y, d, color) {
  const glyphs = ['f', '\u25B6', 'in'];
  glyphs.forEach((g, i) => {
    const cx = x + i * d * 1.3;
    sh(s, 'ellipse', cx, y, d, d, { fill: { color, transparency: 19 } });
    s.addText(g, { x: cx, y, w: d, h: d, align: 'center', valign: 'middle',
      fontFace: PP, fontSize: Math.max(6, Math.round(d * 34)), bold: true, color: W, margin: 0 });
  });
}

/** pill button with a circular arrow badge on its right end */
function button(s, x, y, w, label, fill, labelColor, arrowColor) {
  sh(s, 'roundRect', x + 0.584, y, w, 0.49, { fill, rectRadius: 0.245 });
  sh(s, 'roundRect', x, y, w, 0.49, { fill, rectRadius: 0.245 });
  s.addText(label, { x, y, w, h: 0.49, align: 'center', valign: 'middle',
    fontFace: WS, fontSize: 12, bold: true, charSpacing: 1, color: labelColor, margin: 0 });
  arrowBadge(s, x + w + 0.103, y + 0.085, 0.311, W, arrowColor);
}

/** white disc holding a small coloured arrow */
function arrowBadge(s, x, y, d, discColor, arrowColor) {
  sh(s, 'ellipse', x, y, d, d, { fill: { color: discColor } });
  sh(s, 'rightArrow', x + d * 0.32, y + d * 0.26, d * 0.48, d * 0.48, { fill: { color: arrowColor } });
}

/** skill meter: grey track with a coloured fill on top */
function bar(s, x, y, track, done, color) {
  line(s, x + 0.299, y, track - 0.299, SILVER, 6);
  line(s, x, y, done, color, 6);
}

/* ------------------------------------------------- recurring body copy */
const L1 =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'excepturi sint occaecati cupiditate non provident, similique molesti sunt in culpa qui ' +
  'officia deserunt mollitia animi, id est laborum et et accusamus et iusto idio ' +
  'disginussmis.';
const L2 =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'excepturi sint occaecati cupiditate non provident, similique molesti sunt in culpa qui ' +
  'officia deserunt mollitia';
const L3 =
  'PLACEHOLDER' +
  'PLACEHOLDER';
const L4 =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'occaecati cupiditate non provident, similique molesti sunt in culpa qui officia deserunt ' +
  'mollitia animi, id est laborum et et accusamus et iusto idio';
const L5 =
  'PLACEHOLDER' +
  'praesentium voluptatum.';
const L6 =
  'PLACEHOLDER' +
  'PLACEHOLDER';
const L7 =
  'PLACEHOLDER' +
  'voluptatum deleniti';
const L8 =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'occaecati cupiditate.';
const L9 = 'At vero eos et accusamus et iusto odio dignissimos ducimus qui.';
const L10 =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'excepturi sint occaecati cupiditate non provident, similique molesti.';
const L11 =
  'PLACEHOLDER' +
  'praesentium voluptatum deleniti atque.';
const L12 =
  'PLACEHOLDER' +
  'quos.';
const L13 = 'At vero eos et accusamus eti iusto odio eti ios dignissimos ducimus qui blanditiis ise.';
const L14 =
  'PLACEHOLDER' +
  'PLACEHOLDER' +
  'excepturi sint.';
const L15 = 'At vero eos et accusamus eti iusto odio ise dignissimos ducimus qui blanditiis ise.';
const L16 =
  'When you listen to yourself, everything comes naturally. It comes from inside, like a ' +
  'kind of will to do something. Try to be sensitive. That is yoga';
const L17 =
  'PLACEHOLDER' +
  'praesentium voluptatum deleniti atque corrupti quos dolores et quas et.';
const L18 =
  'PLACEHOLDER' +
  'praesentium voluptatum deleniti atque corrupti quos dolores et quas.';
const L19 =
  'PLACEHOLDER' +
  'praesentium voluptatum deleniti atque corrupti quos dolores.';
const L20 =
  'PLACEHOLDER' +
  'praesentium voluptatum deleniti atque corrupti quos.';

/* ---------------------------------------------------------------- slides */

/** Slide 1 — cover — light wordmark */
function slide1(s) {
  stamp(s,9.48,6.96,W);
  hero(s,0.5,1.82,6.77,1.58,'YOGATHY',BLACK,'center');
  logo(s,0.3,0.31,GREEN);
  tag(s,0.76,3.22,3.76,'KEEP YOUR BODY HEALTHY',GREEN);
  line(s,0,5.17,2.25,AMBER,6);
  url(s,10.78,0.17,GREEN);
}

/** Slide 2 — cover — green wash */
function slide2(s) {
  fade(s,0,0,13.33,7.5,LIME,0.27,0.52);
  hero(s,3.33,2.35,6.77,1.58,'YOGATHY',W,'center');
  tag(s,3.59,3.75,3.76,'KEEP YOUR BODY HEALTHY',W);
  line(s,0,0.85,2.25,AMBER,6);
  logo(s,12.46,0.31,W);
  stamp(s,9.48,6.96,W);
  url(s,0.43,6.95,W);
}

/** Slide 3 — cover — card on photo */
function slide3(s) {
  sh(s,'roundRect',5.86,2.57,7.09,1.89,{fill:W,shadow:SHADOW_DROP,rectRadius:0.32});
  hero(s,6.02,2.57,6.77,1.58,'YOGATHY',BLACK,'center');
  tag(s,6.28,3.96,3.76,'KEEP YOUR BODY HEALTHY',GREEN);
  logo(s,0.3,0.31,GREEN);
  line(s,11.08,1.7,2.25,AMBER,6);
  stamp(s,0.3,6.96,W);
  url(s,10.78,0.17,GREEN);
}

/** Slide 4 — cover — green column */
function slide4(s) {
  sh(s,'rect',1.8,0,3,7.5,{fill:{color:GREEN,transparency:32}});
  hero(s,2.58,1,6.77,1.58,'YOGATHY',W,'center');
  tag(s,2.84,2.4,3.76,'KEEP YOUR BODY HEALTHY',W);
  stamp(s,9.48,6.96,W);
  url(s,2.45,6.95,W);
  logo(s,12.46,0.31,W);
}

/** Slide 5 — intro — social circles */
function slide5(s) {
  kick(s,5.68,1.09,1.66,'INTRODUCTION',W,'center');
  head(s,4.03,1.39,4.96,0.59,'Welcome to Yogathy ','center',W);
  body(s,2.13,5.31,9.07,0.9,L1,GREY,'center');
  kick(s,5.84,4.88,1.66,'INTRODUCTION',GREEN,'center');
  sh(s,'ellipse',5.24,2.72,0.6,0.6,{fill:{color:GREEN,transparency:19}});
  sh(s,'ellipse',6.37,2.06,0.6,0.6,{fill:{color:GREEN,transparency:19}});
  sh(s,'ellipse',7.5,2.72,0.6,0.6,{fill:{color:GREEN,transparency:19}});
  stamp(s,9.68,6.93,AMBER);
  url(s,0.51,6.87,AMBER);
}

/** Slide 6 — intro — quote card */
function slide6(s) {
  sh(s,'roundRect',0.25,2.44,6.56,2.92,{fill:{color:GREEN,transparency:52},shadow:SHADOW_DROP,rectRadius:0.34});
  head(s,0.45,2.64,6.41,2.54,L16,null,W);
  sh(s,'ellipse',5.48,1.85,0.4,0.4,{fill:{color:GREEN,transparency:40}});
  sh(s,'ellipse',6.14,1.62,0.23,0.23,{fill:{color:GREEN,transparency:40}});
  sh(s,'ellipse',6.58,1.41,0.15,0.15,{fill:{color:GREEN,transparency:40}});
  stamp(s,9.48,6.96,W);
  logo(s,12.46,0.31,W);
  url(s,0.43,6.95,W);
}

/** Slide 7 — intro — welcome band */
function slide7(s) {
  fade(s,0,1.66,13.33,3.81,LIME,0.27,0.52);
  kick(s,5.68,0.52,1.66,'INTRODUCTION',GREEN,'center');
  head(s,4.03,0.82,4.96,0.59,[['Welcome to '],['Yoga',GREEN],['thy ']],'center');
  logo(s,0.3,0.31,GREEN);
  url(s,10.78,0.17,GREEN);
  body(s,4.93,6.22,5.85,0.62,L18,GREY);
  kick(s,4.93,5.92,3.1,'YOGATHY – RELAX YOUR BODY',GREEN);
  button(s,5.49,4.8,1.78,'DETAILS',mid(GOLD,AMBER),W,AMBER);
  line(s,0,6.46,2.25,AMBER,6);
}

/** Slide 8 — about — CEO */
function slide8(s) {
  kick(s,0.85,1.36,2.79,'YOGATHY – HEALTHY BODY',GREEN);
  head(s,0.85,1.61,5.26,1.08,[['Yoga is a mirror to look at '],['ourselves',AMBER],[' from within']]);
  body(s,8.88,5.02,3.28,0.9,L11,GREY);
  kick(s,8.88,4.72,2.19,'YOGATHY – OUR CEO',GREEN);
  body(s,7.41,2.23,4.75,1.74,L1,GREY);
  button(s,9.9,6.23,1.78,'DETAILS',mid(GOLD,AMBER),W,AMBER);
  kick(s,7.41,1.95,1.66,'INTRODUCTION',GREEN);
  url(s,10.78,0.17,GREEN);
  line(s,0,3.02,2.25,AMBER,6);
}

/** Slide 9 — about — skill meters */
function slide9(s) {
  // layout 10 furniture
  sh(s,'roundRect',5.6,1.16,7.74,6.34,{fill:W,shadow:{type:'outer',blur:34,offset:3,angle:180,color:BLACK,opacity:0.2},rectRadius:1.06});
  kick(s,0.68,1.28,2.79,'YOGATHY – HEALTHY BODY',GREEN);
  head(s,0.68,1.53,5.26,1.08,[['Yoga is a mirror to look at '],['ourselves',AMBER],[' from within']]);
  body(s,0.68,2.64,4.48,2.01,L1,GREY);
  url(s,10.78,0.17,W);
  logo(s,0.3,0.31,GREEN);
  stamp(s,0.3,6.96,AMBER);
  navlink(s,10.88,0.38,0.59,'Yoga',AMBER);
  navlink(s,9.55,0.38,1.23,'Consultancy',AMBER);
  navlink(s,8.35,0.38,1.09,'Meditation',AMBER);
  bar(s,0.82,5.34,3.32,3.15,GREEN);
  kick(s,0.75,4.94,1.28,'BREATHING',GREY);
  kick14(s,3.81,4.91,0.64,'99%',GREY);
  bar(s,0.82,6.04,3.32,2.54,GREEN);
  kick(s,0.75,5.64,1.39,'MEDITATION',GREY);
  kick14(s,3.81,5.61,0.65,'88%',GREY);
  line(s,6.83,1.29,2.25,AMBER,6);
  sh(s,'ellipse',7.93,2.19,0.84,0.84,{fill:mid(LIME,GREEN)});
  sh(s,'ellipse',11.88,5.49,0.42,0.42,{fill:mid(LIME,GREEN)});
  sh(s,'ellipse',7.25,6.63,0.33,0.33,{fill:mid(LIME,GREEN)});
}

/** Slide 10 — about — diagonal wash */
function slide10(s) {
  fadeQuad(s,7.24,0,6.09,7.5,[{x:1.52},{x:6.09},{x:4.57},{x:0}],GREEN,0.35,1);
  line(s,6.19,1.48,2.25,AMBER,6);
  kick(s,0.98,3.19,1.83,'ABOUT YOGATHY',AMBER);
  head(s,0.98,2.11,5.9,1.08,'The yoga pose you avoid the most you need the most',null,GOLD);
  body(s,0.98,3.49,5.68,1.46,L1,GREY);
  logo(s,0.3,0.31,GREEN);
  url(s,10.78,0.17,AMBER);
  button(s,4.3,5.25,1.78,'READ MORE',mid(GOLD,AMBER),W,AMBER);
  navlink(s,4.43,6.96,0.59,'Yoga',AMBER);
  navlink(s,3.11,6.96,1.23,'Consultancy',AMBER);
  navlink(s,1.91,6.96,1.09,'Meditation',AMBER);
}

/** Slide 11 — about — stats card */
function slide11(s) {
  sh(s,'roundRect',0.83,0.77,11.68,6.13,{fill:W,shadow:SHADOW_WIDE,rectRadius:1.02});
  sh(s,'ellipse',1.83,4.62,0.55,0.55,{fill:mid(LIME,GREEN)});
  sh(s,'ellipse',4.53,1.4,0.4,0.4,{fill:mid(LIME,GREEN)});
  sh(s,'ellipse',3.61,4.68,1.17,1.17,{fill:mid(LIME,GREEN)});
  sh(s,'ellipse',2.03,1.6,2.89,2.89,{fill:{color:ASH,transparency:70}});
  kick(s,5.55,1.54,1.83,'ABOUT YOGATHY',AMBER);
  head(s,5.55,1.84,5.14,1.08,'Yoga takes you into the present moment',null,GOLD);
  body(s,5.55,3.02,5.68,1.46,L1,GREY);
  num44(s,5.62,5.12,1.26,0.84,'265',AMBER);
  kick(s,5.7,5.94,1.11,'MEMBERS',GREY,'center');
  num44(s,7.76,5.12,1.33,0.84,'108',AMBER);
  kick(s,7.83,5.94,1.19,'LOCATION',GREY,'center');
  num44(s,9.89,5.12,1.34,0.84,'158',AMBER);
  kick(s,10.11,5.94,0.91,'WORKS',GREY,'center');
  url(s,1.58,6.24,GREEN);
}

/** Slide 12 — about — tablet illustration */
function slide12(s) {
  sh(s,'roundRect',5.6,0,7.74,6.34,{fill:W,shadow:{type:'outer',blur:34,offset:3,angle:180,color:BLACK,opacity:0.2},flipV:true,rectRadius:1.06});
  sh(s,'ellipse',4.29,2.84,1.13,1.5,{fill:LIME});
  sh(s,'ellipse',4.29,3.63,0.88,0.71,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',1.93,5.01,0.29,0.26,{fill:NAVY});
  sh(s,'roundRect',1.88,4.78,0.29,0.38,{fill:NAVY,rotate:-13.72,rectRadius:0.05});
  sh(s,'roundRect',1.83,4.56,0.29,0.38,{fill:NAVY,rotate:-13.72,rectRadius:0.05});
  sh(s,'roundRect',1.77,4.34,0.29,0.38,{fill:NAVY,rotate:-13.72,rectRadius:0.05});
  sh(s,'roundRect',1.72,4.11,0.29,0.38,{fill:NAVY,rotate:-13.72,rectRadius:0.05});
  sh(s,'ellipse',0.69,1.67,1.78,3.03,{fill:LIME});
  sh(s,'roundRect',0.92,3.32,1.56,1.39,{fill:{color:BLACK,transparency:90},rectRadius:0.23});
  sh(s,'roundRect',0.73,5.11,6.02,0.84,{fill:{color:STONE,transparency:90},rectRadius:0.14});
  sh(s,'roundRect',1.49,1.45,3.26,3.92,{fill:SLATE,rectRadius:0.54});
  sh(s,'custGeom',1.77,1.85,2.83,3.13,{fill:{color:AMBER,transparency:90},points:[{x:0,y:0},{x:2.83,y:0},{x:2.83,y:3.13},{x:0,y:3.13}]});
  sh(s,'roundRect',3.15,5.07,0.23,0.23,{fill:W,rectRadius:0.04});
  sh(s,'custGeom',2.43,3.85,1.38,0.01,{fill:{color:VIOLET,transparency:50},line:{color:AMBER,width:0.39},points:[{x:0,y:0},{x:1.38,y:0},{x:1.38,y:0.01},{x:0,y:0.01}]});
  sh(s,'custGeom',2.43,3.85,0.42,0.01,{fill:W,points:[{x:0,y:0},{x:0.42,y:0},{x:0.42,y:0.01},{x:0,y:0.01}]});
  sh(s,'roundRect',2.43,3.85,0.42,0.01,{fill:W,rectRadius:0});
  sh(s,'roundRect',2.21,2.4,0.53,0.53,{fill:AMBER,rectRadius:0.09});
  sh(s,'roundRect',3.49,2.4,0.53,0.53,{fill:AMBER,rectRadius:0.09});
  sh(s,'ellipse',2.42,3.17,1.39,0.21,{fill:W});
  sh(s,'ellipse',4.88,2.35,0.4,0.39,{fill:ROSE});
  sh(s,'ellipse',5.28,1.97,0.31,0.3,{fill:ROSE});
  sh(s,'ellipse',5.67,2.4,0.31,0.3,{fill:ROSE});
  sh(s,'ellipse',4.35,3.72,2.08,2.1,{fill:GREY});
  sh(s,'ellipse',4.62,4.35,0.91,0.63,{fill:SKIN});
  sh(s,'ellipse',5.04,3.74,1.38,0.6,{fill:CHAR});
  sh(s,'ellipse',5.04,3.77,1.38,0.58,{fill:{color:BLACK,transparency:90}});
  sh(s,'roundRect',4.83,3.94,0.5,0.5,{fill:CHAR,rectRadius:0.08});
  sh(s,'ellipse',4.83,4.16,0.5,0.28,{fill:{color:BLACK,transparency:90}});
  sh(s,'roundRect',4.98,4.35,0.21,0.15,{fill:{color:BLACK,transparency:90},rectRadius:0.03});
  sh(s,'roundRect',4.9,4.12,0.37,0.37,{fill:SKIN,rectRadius:0.06});
  sh(s,'roundRect',4.35,5.46,1.58,0.34,{fill:CHAR,rectRadius:0.06});
  sh(s,'ellipse',4.75,4.57,0.72,1.11,{fill:GOLD});
  sh(s,'ellipse',5.33,4.95,0.06,0.26,{fill:BLACK});
  sh(s,'ellipse',4.75,5.42,0.72,0.26,{fill:BLACK});
  sh(s,'ellipse',4.55,4.85,0.43,0.43,{fill:SKIN});
  sh(s,'ellipse',5.2,4.72,0.45,0.53,{fill:SKIN});
  sh(s,'ellipse',4.92,4.83,0.17,0.34,{fill:SKIN});
  sh(s,'ellipse',5.1,4.83,0.17,0.32,{fill:SKIN});
  sh(s,'roundRect',5.12,5.63,0.33,0.15,{fill:SKIN,rectRadius:0.02});
  sh(s,'roundRect',5.12,5.63,0.33,0.15,{fill:{color:BLACK,transparency:90},rectRadius:0.02});
  sh(s,'roundRect',5.36,5.55,0.46,0.24,{fill:{color:BLACK,transparency:90},rectRadius:0.04});
  sh(s,'roundRect',5.37,5.56,0.53,0.24,{fill:CHAR,rectRadius:0.04});
  sh(s,'ellipse',4.84,5.63,0.75,0.19,{fill:SKIN});
  sh(s,'ellipse',4.52,5.51,0.39,0.28,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',4.4,5.53,0.51,0.28,{fill:CHAR});
  sh(s,'roundRect',4.99,3.94,0.19,0.4,{fill:CHAR,rotate:-83.23,rectRadius:0.03});
  sh(s,'ellipse',4.83,3.98,0.5,0.18,{fill:VIOLET});
  sh(s,'ellipse',4.83,4.03,0.5,0.14,{fill:{color:BLACK,transparency:90}});
  sh(s,'roundRect',6.09,5.48,0.39,0.07,{fill:VIOLET,rectRadius:0.01});
  sh(s,'roundRect',6.17,4.87,0.22,0.48,{fill:GREEN,rectRadius:0.04});
  sh(s,'roundRect',6.17,5.11,0.22,0.24,{fill:GREEN,rectRadius:0.04});
  kick(s,7.13,1.69,1.13,'YOGATHY',AMBER);
  body(s,7.13,3.17,5.68,1.46,L1,GREY);
  head(s,7.13,2.01,5.26,1.08,[['Yoga is a mirror to look at '],['ourselves',AMBER],[' from within']]);
  sh(s,'roundRect',5.88,5.55,5.96,1.41,{fill:GREEN,shadow:SHADOW_DROP,rectRadius:0.24});
  body(s,6.16,6.07,5.47,0.66,L20,W);
  kick(s,6.16,5.74,2.4,'BEST WAY MEDITATION',W);
  url(s,10.78,0.17,AMBER);
}

/** Slide 13 — about — stretch illustration */
function slide13(s) {
  sh(s,'ellipse',7.57,1.24,4.96,1.18,{fill:{color:AMBER,transparency:90}});
  sh(s,'ellipse',7.13,2.41,5.5,2.54,{fill:{color:AMBER,transparency:90}});
  sh(s,'roundRect',7.27,4.75,4.2,0.4,{fill:LIME,rectRadius:0.07});
  sh(s,'ellipse',7.73,3.08,3.49,1.95,{fill:GREY});
  sh(s,'ellipse',8.92,3.1,1.95,0.56,{fill:'FFC1C7'});
  sh(s,'ellipse',8.92,3.1,1.71,0.56,{fill:BLACK});
  sh(s,'ellipse',10.53,3.37,0.16,0.22,{fill:BLACK});
  sh(s,'ellipse',9.95,3.61,0.25,0.94,{fill:'965D7B'});
  sh(s,'ellipse',9.95,4.46,0.23,0.1,{fill:BLACK});
  sh(s,'ellipse',9.97,4.11,0.04,0.28,{fill:BLACK});
  sh(s,'ellipse',8.73,4.43,2.46,0.58,{fill:'FFC1C7'});
  sh(s,'ellipse',8.73,4.48,0.39,0.32,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.28,4.32,0.83,0.48,{fill:'565988'});
  sh(s,'ellipse',9.04,3.54,0.65,0.58,{fill:'FFC1C7'});
  sh(s,'ellipse',9.54,3.54,0.15,0.22,{fill:{color:BLACK,transparency:90}});
  sh(s,'roundRect',9.55,3.38,0.43,0.43,{fill:'FFC1C7',rectRadius:0.07});
  sh(s,'ellipse',9.77,4.55,1.16,0.19,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',9.27,3.83,1.66,0.9,{fill:'FFC1C7'});
  sh(s,'ellipse',8.79,3.52,0.85,0.71,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.79,3.51,0.85,0.71,{fill:'565988'});
  sh(s,'ellipse',8.33,4.32,0.77,0.17,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.34,3.51,1.29,0.93,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.92,3.5,0.54,0.15,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.33,3.5,1.3,0.97,{fill:'E8F4FF'});
  sh(s,'ellipse',8.59,4.04,0.27,0.33,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.61,4.3,0.25,0.11,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.49,4.51,0.25,0.11,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.53,4.59,0.25,0.11,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.56,4.47,0.25,0.11,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',9.73,3.38,0.25,0.42,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',9.74,3.36,0.36,0.48,{fill:'965D7B'});
  sh(s,'ellipse',8.74,4.68,0.71,0.23,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',7.76,4.43,1.69,0.52,{fill:'FFC1C7'});
  sh(s,'ellipse',8.28,4.48,0.35,0.24,{fill:{color:BLACK,transparency:90}});
  sh(s,'ellipse',8.28,4.47,0.35,0.24,{fill:'565988'});
  sh(s,'ellipse',8.1,4.54,0.48,0.16,{fill:W,line:{color:BLACK,width:0.36}});
  sh(s,'ellipse',9.74,3.36,0.17,0.48,{fill:{color:BLACK,transparency:90}});
  sh(s,'roundRect',11.32,4.77,0.44,0.06,{fill:{color:VIOLET,transparency:90},rectRadius:0.01});
  sh(s,'ellipse',11.43,4.66,0.22,0.13,{fill:LIME});
  sh(s,'roundRect',12.02,4.9,0.44,0.06,{fill:{color:VIOLET,transparency:90},rectRadius:0.01});
  sh(s,'ellipse',12.13,4.79,0.22,0.13,{fill:LIME});
  sh(s,'roundRect',11.58,5.17,0.44,0.06,{fill:{color:VIOLET,transparency:90},rectRadius:0.01});
  sh(s,'ellipse',11.69,5.06,0.22,0.13,{fill:LIME});
  sh(s,'roundRect',7.13,4.72,0.44,0.06,{fill:{color:VIOLET,transparency:90},rectRadius:0.01});
  sh(s,'ellipse',7.24,4.61,0.22,0.13,{fill:LIME});
  sh(s,'roundRect',6.91,5.14,0.44,0.06,{fill:{color:VIOLET,transparency:90},rectRadius:0.01});
  sh(s,'ellipse',7.02,5.04,0.22,0.13,{fill:LIME});
  kick(s,0.91,1.38,1.13,'YOGATHY',AMBER);
  head(s,0.91,1.69,5.14,1.08,'Yoga takes you into the present moment',null,GOLD);
  sh(s,'roundRect',1.02,5.42,3.71,1.51,{fill:W,shadow:SHADOW_WIDE,rectRadius:0.25});
  body(s,1.25,5.85,3.28,0.9,L11,GREY);
  kick(s,1.25,5.55,1.43,'WHO WE ARE',GREEN);
  sh(s,'roundRect',5.04,5.38,3.71,1.55,{fill:W,shadow:SHADOW_WIDE,rectRadius:0.26});
  body(s,5.27,5.81,3.28,0.9,L11,GREY);
  kick(s,5.27,5.51,1.45,'WHAT WE DO',GREEN);
  sh(s,'ellipse',4.25,4.57,1.17,1.17,{fill:mid(LIME,GREEN)});
  body(s,0.91,2.87,5.68,1.46,L1,GREY);
  line(s,11.08,6.29,2.25,AMBER,6);
  url(s,10.78,0.17,AMBER);
  logo(s,0.3,0.31,GREEN);
  sh(s,'ellipse',8.48,6.72,0.44,0.44,{fill:mid(LIME,GREEN)});
  sh(s,'ellipse',0.91,6.72,0.44,0.44,{fill:mid(LIME,GREEN)});
  sh(s,'ellipse',4.67,4.67,0.74,0.91,{fill:W});
}

/** Slide 14 — how it works — card + strips */
function slide14(s) {
  fade(s,8.68,4.93,4.66,2.57,LIME,0.47,1);
  fade(s,8.68,2.34,4.66,2.57,LIME,0.47,1);
  fade(s,8.68,-0.01,4.66,2.6,LIME,0.47,1);
  sh(s,'roundRect',0.98,1.82,7.06,4.17,{fill:W,shadow:SHADOW_WIDE,rectRadius:0.69});
  body(s,8.95,4.17,4.08,0.62,L12,W);
  body(s,8.95,1.82,4.08,0.62,L12,W);
  body(s,8.95,6.63,4.08,0.62,L12,W);
  kick(s,1.71,2.04,2.58,'HOW – YOGATHY WORKS',GREEN);
  body(s,1.71,3.52,5.68,1.46,L1,GREY);
  head(s,1.71,2.36,5.26,1.08,[['Yoga is a mirror to look at '],['ourselves',AMBER],[' from within']]);
  line(s,0,0.89,2.25,AMBER,6);
  button(s,5.03,5.23,1.78,'READ MORE',mid(GOLD,AMBER),W,AMBER);
  url(s,0.25,6.91,AMBER);
}

/** Slide 15 — how it works — single feature card */
function slide15(s) {
  // layout 14 furniture
  sh(s,'rect',5.11,0,8.22,7.5,{fill:{color:GREEN,transparency:91}});
  kick(s,0.53,1.14,1.83,'ABOUT YOGATHY',AMBER);
  head(s,0.53,1.44,6.14,1.08,'The yoga pose you avoid the most you need the most',null,GOLD);
  body(s,0.53,2.63,4.37,2.01,L1,GREY);
  sh(s,'roundRect',5.6,4.02,2.7,3.06,{fill:GREEN,rectRadius:0.45});
  sh(s,'ellipse',6.85,4.24,0.74,0.91,{fill:W});
  body(s,5.81,5.52,2.29,1.18,L5,W,'center');
  kick(s,6.39,5.26,1.13,'YOGATHY',W,'center');
  line(s,0,6.18,2.25,AMBER,6);
}

/** Slide 16 — how it works — three feature cards */
function slide16(s) {
  sh(s,'roundRect',3.97,4.05,2.7,3.06,{fill:GREEN,rectRadius:0.45});
  sh(s,'ellipse',5.21,4.27,0.74,0.91,{fill:W});
  body(s,4.17,5.55,2.29,1.18,L5,W,'center');
  kick(s,4.6,5.29,1.43,'WHO WE ARE',W,'center');
  sh(s,'roundRect',9.77,4.03,2.7,3.06,{fill:GREEN,rectRadius:0.45});
  body(s,9.98,5.53,2.29,1.18,L5,W,'center');
  kick(s,10.23,5.28,1.77,'HOW WE WORKS',W,'center');
  sh(s,'ellipse',10.76,4.68,0.82,0.48,{fill:W});
  sh(s,'roundRect',6.87,4.02,2.7,3.06,{fill:GREEN,rectRadius:0.45});
  body(s,7.08,5.52,2.29,1.18,L5,W,'center');
  kick(s,7.49,5.26,1.45,'WHAT WE DO',W,'center');
  sh(s,'roundRect',7.8,4.54,0.84,0.7,{fill:W,rectRadius:0.12});
  kick(s,5.49,0.72,2.58,'HOW – YOGATHY WORKS',GREEN);
  body(s,5.49,2.2,5.68,1.46,L1,GREY);
  head(s,5.49,1.05,5.26,1.08,[['Yoga is a mirror to look at '],['ourselves',AMBER],[' from within']]);
  arrowBadge(s,12.27,5.47,0.41,W,AMBER);
  arrowBadge(s,3.77,5.47,0.41,W,AMBER);
  url(s,0.25,6.91,AMBER);
}

/** Slide 17 — team — four name plates */
function slide17(s) {
  // layout 16 furniture
  sh(s,'rect',0,0,8.22,7.5,{fill:{color:GREEN,transparency:91}});
  sh(s,'roundRect',6.67,2.62,2.7,2.03,{fill:GREEN,rectRadius:0.34});
  sh(s,'roundRect',9.61,4.18,2.7,1.08,{fill:GREEN,rectRadius:0.18});
  name(s,9.8,4.74,2.33,0.4,'Demi Watson',W,'center');
  sh(s,'roundRect',3.74,4.18,2.7,1.08,{fill:GREEN,rectRadius:0.18});
  name(s,3.93,4.74,2.33,0.4,'Shopia Martino',W,'center');
  sh(s,'roundRect',0.84,4.18,2.7,1.08,{fill:GREEN,rectRadius:0.18});
  kick(s,2.41,0.67,1.82,'YOGATHY TEAMS',GREEN);
  head(s,2.41,0.99,5.26,1.08,[['Meets '],['yogathy',AMBER],[' professional teams']]);
  body(s,6.88,3.12,2.29,1.18,L5,W,'center');
  kick(s,7.38,2.87,1.29,'OUR TEAMS',W,'center');
  name(s,1.02,4.74,2.33,0.4,'Oliver Cooper',W,'center');
  line(s,-0.01,2.25,2.25,AMBER,6);
  body(s,2.46,5.82,7.27,1.18,L1,GREY);
  url(s,10.78,0.17,AMBER);
  button(s,6.84,4.82,1.78,'READ MORE',mid(LIME,GREEN),W,LIME);
  bar(s,7.64,1.88,3.32,2.98,GREEN);
  kick(s,7.57,1.48,1.48,'TEAM WORKS',GREY);
  kick14(s,10.63,1.45,0.65,'98%',GREY);
}

/** Slide 18 — team — trainer grid */
function slide18(s) {
  social(s,11.04,0.48,0.43,GREEN);
  name(s,3.52,2.81,1.94,0.4,'Eva Mattheus',GREEN);
  body(s,3.52,3.21,2.83,0.62,L9,GREY);
  name(s,3.52,5.53,1.77,0.4,'Naomi Scott',GREEN);
  body(s,3.52,5.93,2.83,0.62,L9,GREY);
  name(s,9.22,2.81,1.79,0.4,'Rosie James',GREEN,'center');
  body(s,9.3,3.21,2.83,0.62,L9,GREY);
  name(s,9.3,5.53,1.62,0.4,'Keira Davis',GREEN,'center');
  body(s,9.3,5.93,2.83,0.62,L9,GREY);
  kick(s,2.17,0.73,1.98,'TRAINER YOGATHY',AMBER);
  head(s,2.17,1.04,6.14,0.59,'Yogathy expert teams',null,GOLD);
  line(s,0,1.75,2.25,AMBER,6);
  url(s,11.43,6.87,AMBER);
}

/** Slide 19 — team — four columns */
function slide19(s) {
  kick(s,5.76,0.71,1.82,'YOGATHY TEAMS',GREEN,'center');
  head(s,4.04,1.03,5.26,1.08,[['Meets '],['yogathy',AMBER],[' professional teams']],'center');
  name(s,0.96,2.81,1.62,0.4,'Baeily Reid',GREEN,'center');
  body(s,0.35,3.21,2.83,0.62,L9,GREY,'center');
  name(s,3.83,2.81,2.38,0.4,'William Morton',GREEN);
  body(s,3.62,3.21,2.83,0.62,L9,GREY,'center');
  name(s,7.45,2.81,1.66,0.4,'Katie Owen',GREEN,'center');
  body(s,6.86,3.21,2.83,0.62,L9,GREY,'center');
  name(s,10.66,2.81,1.72,0.4,'Daisy Smith',GREEN,'center');
  body(s,10.11,3.21,2.83,0.62,L9,GREY,'center');
  logo(s,0.3,0.31,GREEN);
  url(s,11.28,0.18,AMBER);
}

/** Slide 20 — team — two trainer cards */
function slide20(s) {
  // layout 19 furniture
  sh(s,'rect',3.45,1.96,2.95,4.53,{fill:GREEN});
  sh(s,'rect',6.93,1.96,2.95,4.53,{fill:W,shadow:SHADOW});
  name(s,3.63,2.8,1.79,0.4,'Rosie James',W,'center');
  kick(s,3.63,2.51,1,'TRAINER',W);
  name(s,7.19,2.74,1.77,0.4,'Naomi Scott',BLACK,'center');
  kick(s,7.18,2.45,1,'TRAINER',GREEN);
  bar(s,3.69,4.76,2.42,1.69,AMBER);
  kick(s,3.63,4.36,1.39,'MEDITATION',W);
  bar(s,3.69,5.39,2.42,1.69,AMBER);
  kick(s,3.63,4.99,1.28,'BREATHING',W);
  bar(s,3.69,6.11,2.42,1.69,AMBER);
  kick(s,3.63,5.7,1.44,'METABOLISM',W);
  bar(s,7.25,4.76,2.42,1.69,GREEN);
  kick(s,7.18,4.36,1.39,'MEDITATION',GREY);
  bar(s,7.25,5.39,2.42,1.69,GREEN);
  kick(s,7.18,4.99,1.28,'BREATHING',GREY);
  bar(s,7.25,6.11,2.42,1.69,GREEN);
  kick(s,7.18,5.7,1.44,'METABOLISM',GREY);
  body(s,3.63,3.18,2.49,0.9,L15,W);
  body(s,7.19,3.18,2.49,0.9,L15,GREY);
  social(s,7.25,2.15,0.21,GREEN);
  social(s,3.71,2.15,0.21,AMBER);
  kick(s,2.17,0.66,1.98,'TRAINER YOGATHY',AMBER);
  head(s,2.17,0.96,6.14,0.59,'Yogathy expert teams',null,GOLD);
  line(s,11.08,1.37,2.25,AMBER,6);
  stamp(s,9.68,6.93,AMBER);
  url(s,0.51,6.87,AMBER);
}

/** Slide 21 — team — quote card */
function slide21(s) {
  sh(s,'roundRect',9.55,1.13,2.95,3.06,{fill:W,shadow:SHADOW,rectRadius:0.38});
  kick(s,3.89,1.12,2.09,'MEMBERS YOGATHY',AMBER);
  head(s,3.89,1.42,6.14,0.59,'Yogathy all expert teams',null,GOLD);
  url(s,11.01,6.94,AMBER);
  body(s,3.89,2.26,4.86,1.74,L1,GREY);
  button(s,9.8,3.39,1.78,'READ MORE',mid(LIME,GREEN),W,LIME);
  body(s,9.88,1.92,2.29,1.18,L5,GREY,'center');
  kick(s,10.38,1.55,1.29,'OUR TEAMS',GREEN,'center');
  stamp(s,9.54,0.24,AMBER);
}

/** Slide 22 — team — carousel */
function slide22(s) {
  // layout 21 furniture
  sh(s,'roundRect',0.43,2.48,2.95,3.36,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.41});
  sh(s,'roundRect',3.55,2.48,2.95,3.36,{fill:W,shadow:SHADOW_GREY,rectRadius:0.41});
  sh(s,'roundRect',6.67,2.48,2.95,3.36,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.41});
  sh(s,'roundRect',9.78,2.48,2.95,3.36,{fill:W,shadow:SHADOW_GREY,rectRadius:0.41});
  kick(s,5.76,0.71,1.82,'YOGATHY TEAMS',GREEN,'center');
  head(s,4.04,1.03,5.26,1.08,[['Meets '],['yogathy',AMBER],[' professional teams']],'center');
  name(s,0.94,5.05,1.94,0.4,'Eva Mattheus',W,'center');
  logo(s,0.3,0.31,GREEN);
  url(s,11.28,0.18,AMBER);
  name(s,3.97,5.05,2.11,0.4,'Joshua White',GREEN,'center');
  name(s,7.2,5.05,1.92,0.4,'Demi Watson',W,'center');
  name(s,10.38,5.05,1.79,0.4,'Rosie James',GREEN,'center');
  arrowBadge(s,12.54,3.75,0.58,GREEN,W);
  arrowBadge(s,0.14,3.75,0.58,W,GREEN);
  stamp(s,9.68,6.93,AMBER);
}

/** Slide 23 — team — member grid */
function slide23(s) {
  kick(s,1.99,0.41,2.09,'MEMBERS YOGATHY',AMBER);
  head(s,1.99,0.71,6.14,0.59,'Yogathy all expert teams',null,GOLD);
  stamp(s,9.54,0.24,AMBER);
  name(s,3.73,2.37,2.09,0.4,'Luciana Mary',GREEN,'center');
  body(s,3.65,2.77,2.83,0.62,L9,GREY);
  name(s,9.88,2.37,1.67,0.4,'Laura Diaz',GREEN,'center');
  body(s,9.86,2.77,2.83,0.62,L9,GREY);
  name(s,3.64,5.25,1.8,0.4,'Samuel Lee',GREEN,'center');
  body(s,3.65,5.65,2.83,0.62,L9,GREY);
  name(s,9.86,5.25,2.31,0.4,'Cecilia Moreno',GREEN);
  body(s,9.86,5.65,2.83,0.62,L9,GREY);
  social(s,3.78,3.51,0.31,GREEN);
  social(s,3.78,6.37,0.31,GREEN);
  social(s,9.97,3.51,0.31,GREEN);
  social(s,9.97,6.37,0.31,GREEN);
}

/** Slide 24 — service — vertical band */
function slide24(s) {
  fade(s,3.59,0.01,3.8,7.49,SILVER,0.31,1,'h');
  sh(s,'ellipse',7.98,3.7,0.69,0.75,{fill:W});
  sh(s,'ellipse',8.03,5.77,0.45,0.94,{fill:W,rotate:45});
  sh(s,'ellipse',7.98,1.66,0.77,0.76,{fill:W});
  kick(s,1.06,1.75,1.97,'YOGATHY SERVICE',GREEN);
  head(s,1.06,2.08,4.59,1.08,[['Service of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,1.06,3.27,4.86,1.74,L1,GREY);
  body(s,9.13,1.73,3.14,0.9,L5,W);
  kick(s,9.13,1.43,1.56,'FIRST SERVICE',W);
  body(s,9.13,3.71,3.14,0.9,L5,W);
  kick(s,9.13,3.4,1.82,'SECOND SERVICE',W);
  body(s,9.13,5.77,3.14,0.9,L5,W);
  kick(s,9.13,5.47,1.61,'THIRD SERVICE',W);
  url(s,0.51,6.87,AMBER);
  logo(s,0.3,0.31,GREEN);
  line(s,0.01,5.7,2.25,AMBER,6);
}

/** Slide 25 — service — four icon cards */
function slide25(s) {
  sh(s,'roundRect',0.45,3.41,2.95,3.11,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.41});
  sh(s,'roundRect',3.6,3.41,2.95,3.11,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.41});
  sh(s,'roundRect',6.76,3.41,2.95,3.11,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.41});
  sh(s,'roundRect',9.91,3.41,2.95,3.11,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.41});
  kick(s,4.05,0.55,2.02,'SERVICE  YOGATHY',AMBER);
  head(s,4.05,0.85,6.14,0.59,'Yogathy all best service',null,GOLD);
  body(s,0.78,4.85,2.29,1.18,L5,W,'center');
  kick(s,1.28,4.59,1.28,'ANNOUNCE',W,'center');
  sh(s,'ellipse',4.73,3.62,0.69,0.75,{fill:W});
  sh(s,'ellipse',11.16,3.56,0.45,0.94,{fill:W,rotate:45});
  sh(s,'ellipse',1.6,3.58,0.77,0.76,{fill:W});
  sh(s,'ellipse',7.95,3.47,0.57,1.04,{fill:W,rotate:45});
  body(s,3.9,4.85,2.29,1.18,L5,W,'center');
  kick(s,4.72,4.59,0.65,'TIME',W,'center');
  body(s,7.11,4.85,2.29,1.18,L5,W,'center');
  kick(s,7.7,4.59,1.09,'EXERCISE',W,'center');
  body(s,10.25,4.85,2.29,1.18,L5,W,'center');
  kick(s,10.97,4.59,0.84,'NOTED',W,'center');
  stamp(s,9.54,6.88,AMBER);
  url(s,0.51,6.87,AMBER);
  logo(s,0.3,0.31,GREEN);
}

/** Slide 26 — service — three panels */
function slide26(s) {
  sh(s,'rect',1.14,2.85,3.29,3.76,{fill:W,shadow:SHADOW});
  sh(s,'rect',4.98,2.85,3.29,3.76,{fill:GREEN,shadow:SHADOW});
  sh(s,'rect',8.82,2.85,3.29,3.76,{fill:W,shadow:SHADOW});
  kick(s,1.31,0.89,2.02,'SERVICE  YOGATHY',W);
  head(s,1.31,1.2,3.89,1.08,'Yogathy have many best service',null,W);
  body(s,1.34,3.49,2.87,2.85,L4,GREY);
  kick(s,1.34,3.19,1.29,'BREATHING',GREEN);
  body(s,9.07,3.49,2.87,2.85,L4,GREY);
  kick(s,9.07,3.19,1.44,'METABOLISM',GREEN);
  body(s,5.2,3.49,2.87,2.85,L4,W);
  kick(s,5.2,3.19,1.39,'MEDITATION',W);
  url(s,11.23,0.18,AMBER);
  arrowBadge(s,11.36,6.34,0.58,GREEN,W);
  arrowBadge(s,3.63,6.34,0.58,GREEN,W);
  arrowBadge(s,7.51,6.34,0.58,W,GREEN);
}

/** Slide 27 — service — three tall cards */
function slide27(s) {
  sh(s,'roundRect',6.93,2.52,2.95,4.53,{fill:W,shadow:SHADOW,rectRadius:0.41});
  sh(s,'roundRect',3.74,2.52,2.95,4.53,{fill:GREEN,shadow:SHADOW,rectRadius:0.41});
  sh(s,'roundRect',0.55,2.52,2.95,4.53,{fill:W,shadow:SHADOW,rectRadius:0.41});
  kick(s,1.34,0.77,1.97,'YOGATHY SERVICE',GREEN);
  head(s,1.34,1.09,4.59,1.08,[['Service of '],['yogathy',AMBER],[' help and guide you']]);
  kick(s,0.74,5.49,1.29,'BREATHING',GREEN);
  kick(s,7.17,5.49,1.44,'METABOLISM',GREEN);
  kick(s,3.9,5.49,1.39,'MEDITATION',W);
  body(s,0.74,5.81,2.52,0.9,L13,GREY);
  body(s,3.91,5.81,2.52,0.9,L13,W);
  body(s,7.15,5.81,2.52,0.9,L13,GREY);
  url(s,11.23,0.18,W);
  logo(s,0.3,0.31,GREEN);
}

/** Slide 28 — service — three pills */
function slide28(s) {
  sh(s,'roundRect',1.02,6.17,2.31,0.53,{fill:GREEN,shadow:SHADOW,rectRadius:0.07});
  kick(s,1.42,6.28,1.58,'REGISTRATION',W,'center');
  social(s,0.33,0.48,0.43,GREEN);
  sh(s,'roundRect',4.96,6.17,2.31,0.53,{fill:GREEN,shadow:SHADOW,rectRadius:0.07});
  kick(s,5.34,6.28,1.65,'CHOOSE CLASS',W,'center');
  sh(s,'roundRect',9.69,6.17,2.31,0.53,{fill:GREEN,shadow:SHADOW,rectRadius:0.07});
  kick(s,10.21,6.28,1.4,'START YOGA',W,'center');
  kick(s,1.59,1.48,1.97,'YOGATHY SERVICE',GREEN);
  head(s,1.59,1.81,4.59,1.08,[['Service of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,6.41,1.75,5.58,1.18,L2,GREY);
  line(s,11.08,1.51,2.25,AMBER,6);
  url(s,11.23,0.18,AMBER);
}

/** Slide 29 — service — four tiles */
function slide29(s) {
  // layout 28 furniture
  sh(s,'roundRect',0.61,0.36,2.95,2.88,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.4});
  sh(s,'roundRect',4.17,0.36,2.95,2.88,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.4});
  sh(s,'roundRect',0.61,3.88,2.95,2.88,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.4});
  sh(s,'roundRect',4.17,3.88,2.95,2.88,{fill:GREEN,shadow:SHADOW_GREY,rectRadius:0.4});
  kick(s,8.02,1.94,1.97,'YOGATHY SERVICE',GREEN);
  head(s,8.02,2.27,4.59,1.08,[['Service of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,8.02,3.4,4.05,1.74,L2,GREY);
  url(s,11.23,0.18,AMBER);
  line(s,11.08,6.8,2.25,AMBER,6);
  kick(s,1.29,0.53,1.58,'REGISTRATION',W,'center');
  kick(s,4.82,0.53,1.65,'CHOOSE CLASS',W,'center');
  kick(s,1.38,4.07,1.4,'START YOGA',W,'center');
  kick(s,4.73,4.07,1.84,'CHOOSE TRAINER',W,'center');
}

/** Slide 30 — break slide — light */
function slide30(s) {
  hero(s,0.76,1.56,4.79,3.06,[['BREAK',null,{breakLine:true}],['SLIDE']],BLACK,'center');
  tag(s,0.89,4.44,3.76,'KEEP YOUR BODY HEALTHY',AMBER);
  line(s,0,5.17,2.25,GREEN,6);
  url(s,11.23,0.18,AMBER);
  social(s,0.33,0.48,0.43,GREEN);
  stamp(s,9.48,6.96,W);
}

/** Slide 31 — break slide — green wash */
function slide31(s) {
  fade(s,0,0.02,13.33,7.5,LIME,0.27,0.52);
  hero(s,0.88,2.19,4.79,1.58,'BREAK',W,'center');
  tag(s,9.14,6.76,3.76,'KEEP YOUR BODY HEALTHY',W);
  url(s,0.43,6.95,W);
  social(s,11.34,0.48,0.43,GREEN);
  line(s,0,1.46,2.25,W,6);
  button(s,3.93,4.85,1.78,'BREAK NOW',mid(LIME,GREEN),W,LIME);
  hero(s,2.18,3.31,4.22,1.58,'SLIDE',W);
}

/** Slide 32 — break slide — right photo */
function slide32(s) {
  hero(s,7.93,2.31,4.79,1.58,'BREAK',GREEN,'center');
  tag(s,1.79,5.73,3.76,'KEEP YOUR BODY HEALTHY',W,'right');
  hero(s,8.5,3.51,4.22,1.58,'SLIDE',BLACK);
  logo(s,0.3,0.31,GREEN);
  stamp(s,9.54,6.88,AMBER);
  url(s,11.08,0.18,AMBER);
  button(s,2.31,6.22,1.78,'BREAK NOW',mid(LIME,GREEN),W,LIME);
}

/** Slide 33 — break slide — green column */
function slide33(s) {
  sh(s,'rect',8.72,0,3,7.5,{fill:{color:GREEN,transparency:32}});
  hero(s,5.72,1.76,4.79,3.06,[['BREAK',null,{breakLine:true}],['SLIDE']],W,'right');
  tag(s,6.55,4.82,3.76,'KEEP YOUR BODY HEALTHY',W);
  url(s,9.63,6.95,W);
  stamp(s,0.58,6.96,W);
  social(s,0.44,0.48,0.43,GREEN);
}

/** Slide 34 — portfolio — bottom caption */
function slide34(s) {
  kick(s,1.2,5.48,2.24,'YOGATHY PORTFOLIO',GREEN);
  head(s,1.2,5.81,4.59,1.08,[['Portfolio of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,6.03,5.75,5.58,1.18,L2,GREY);
  sh(s,'ellipse',8.74,2.45,0.31,0.31,{fill:{color:GREEN,transparency:19}});
  social(s,11.96,7.01,0.31,GREEN);
  sh(s,'ellipse',4.29,2.45,0.31,0.31,{fill:{color:GREEN,transparency:19}});
}

/** Slide 35 — portfolio — right card */
function slide35(s) {
  sh(s,'roundRect',9.37,2.71,3.18,3.9,{fill:GREEN,shadow:SHADOW,rectRadius:0.44});
  body(s,9.53,3,2.87,2.85,L4,W);
  kick(s,4.22,3.6,2.24,'YOGATHY PORTFOLIO',GREEN);
  head(s,4.22,3.93,4.59,1.08,[['Portfolio of '],['yogathy',AMBER],[' help and guide you']]);
  button(s,9.74,5.94,1.78,'READ MORE',W,GREEN,LIME);
  body(s,4.22,5.1,4.74,1.46,L2,GREY);
  social(s,4.22,7.01,0.31,GREEN);
  url(s,1.03,0.18,AMBER);
}

/** Slide 36 — portfolio — panel */
function slide36(s) {
  sh(s,'rect',7.27,2.16,3.29,3.76,{fill:GREEN,shadow:SHADOW});
  body(s,7.5,2.79,2.87,2.85,L4,W);
  kick(s,7.5,2.49,1.76,'BEST PORTFOLIO',W);
  arrowBadge(s,9.81,5.64,0.58,W,GREEN);
  kick(s,0.48,1.16,2.59,'PORTFOLIO OF  YOGATHY',AMBER);
  head(s,0.48,1.46,3.12,1.08,'Yogathy all best portfolio',null,GOLD);
  body(s,0.48,2.66,3.43,2.01,L2,GREY);
  body(s,0.48,4.7,3.43,1.18,L19,GREY);
  url(s,0.48,6.96,AMBER);
  logo(s,0.3,0.31,GREEN);
}

/** Slide 37 — portfolio — three cards */
function slide37(s) {
  // layout 32 furniture
  sh(s,'roundRect',6.48,2.11,2.94,5.18,{fill:W,shadow:SHADOW,rectRadius:0.22});
  sh(s,'ellipse',7.08,6.04,0.97,0.97,{fill:MIST});
  sh(s,'ellipse',8.05,2.25,0.97,0.97,{fill:MIST});
  sh(s,'roundRect',3.37,2.11,2.94,5.18,{fill:W,shadow:SHADOW,rectRadius:0.22});
  sh(s,'ellipse',3.98,6.04,0.97,0.97,{fill:MIST});
  sh(s,'ellipse',4.95,2.25,0.97,0.97,{fill:MIST});
  sh(s,'roundRect',0.27,2.11,2.94,5.18,{fill:GREEN,shadow:SHADOW,rectRadius:0.22});
  sh(s,'ellipse',0.73,6.04,0.97,0.97,{fill:LIME});
  sh(s,'ellipse',1.7,2.25,0.97,0.97,{fill:LIME});
  line(s,9.18,2.68,2.25,AMBER,6);
  kick(s,2.44,0.93,2.59,'PORTFOLIO OF  YOGATHY',AMBER);
  head(s,2.44,1.23,5.52,0.59,'Yogathy all best portfolio',null,GOLD);
  body(s,9.59,3.21,3.43,2.01,L2,GREY);
  button(s,10.41,5.67,1.78,'MORE DETAILS',mid(LIME,GREEN),W,LIME);
  url(s,3.96,6.96,AMBER);
  logo(s,0.3,0.31,GREEN);
  social(s,11.9,0.21,0.31,GREEN);
}

/** Slide 38 — portfolio — green banner */
function slide38(s) {
  sh(s,'roundRect',8.16,4.8,4.24,2.19,{fill:W,shadow:SHADOW,rectRadius:0.31});
  sh(s,'roundRect',0.12,2.1,4.59,2.69,{fill:mid(LIME,GREEN),shadow:SHADOW,rectRadius:0.38});
  kick(s,0.3,2.63,2.61,'YOGATHY ALL PORTFOLIO',W);
  head(s,0.3,2.95,4.59,1.08,'Portfolio of yogathy help and guide you',null,W);
  url(s,2.82,4.3,W);
  social(s,11.9,3.96,0.31,GREEN);
  sh(s,'ellipse',8.74,1.84,0.31,0.31,{fill:{color:GREEN,transparency:19}});
  body(s,8.26,5.03,4.05,1.74,L2,GREY);
  sh(s,'ellipse',7.77,4.64,0.37,0.38,{fill:W,shadow:SHADOW});
  sh(s,'ellipse',7.4,4.45,0.25,0.25,{fill:W,shadow:SHADOW});
}

/** Slide 39 — portfolio — carousel */
function slide39(s) {
  kick(s,6.84,2.1,2.24,'YOGATHY PORTFOLIO',GREEN);
  head(s,6.84,2.42,4.59,1.08,[['Portfolio of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,6.84,3.59,4.74,1.46,L2,GREY);
  social(s,11.9,0.21,0.31,GREEN);
  button(s,9.54,5.41,1.78,'MORE DETAILS',mid(LIME,GREEN),W,LIME);
  url(s,5.79,6.91,AMBER);
  stamp(s,9.68,6.93,AMBER);
  arrowBadge(s,0.68,5.59,0.58,W,GREEN);
  arrowBadge(s,4.98,1.45,0.58,W,GREEN);
}

/** Slide 40 — portfolio — centred card */
function slide40(s) {
  sh(s,'rect',3.08,2.62,5.81,2.27,{fill:W,shadow:SHADOW});
  kick(s,3.24,3.02,2.59,'PORTFOLIO OF  YOGATHY',AMBER);
  head(s,3.24,3.33,5.52,0.59,'Yogathy all best portfolio',null,GOLD);
  line(s,6.64,4.29,2.25,AMBER,6);
  url(s,3.24,4.45,AMBER);
}

/** Slide 41 — mockup — desktop screens */
function slide41(s) {
  // layout 36 furniture
  sh(s,'rect',0,3.87,13.33,1.54,{fill:GREEN});
  photo(s,8.45,2.86,3.8,3.8);
  photo(s,1.35,2.86,3.8,3.8);
  photo(s,3.81,1.42,5.92,5.92);
  kick(s,1.09,0.51,2.02,'YOGATHY MOCKUP',GREEN);
  head(s,1.09,0.88,4.59,1.08,[['Mockup of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,7.65,0.77,4.74,1.18,L10,GREY);
  url(s,0.64,6.96,AMBER);
  arrowBadge(s,9.25,4.09,0.58,W,GREEN);
  arrowBadge(s,3.65,4.09,0.58,W,GREEN);
}

/** Slide 42 — mockup — phone stack */
function slide42(s) {
  // layout 37 furniture
  sh(s,'rect',0,0,1.05,7.5,{fill:GREEN});
  sh(s,'triangle',0.68,1.62,0.48,0.41,{fill:W,rotate:270});
  sh(s,'triangle',0.68,3.69,0.48,0.41,{fill:W,rotate:270});
  sh(s,'triangle',0.68,5.6,0.48,0.41,{fill:W,rotate:270});
  photo(s,10.67,3.46,2.27,4.64);
  photo(s,8.4,2.15,2.27,4.64);
  photo(s,10.67,-1.19,2.27,4.64);
  body(s,1.71,2.98,5.52,0.9,L8,GREY);
  kick(s,1.71,2.67,1.29,'BREATHING',GREEN);
  kick(s,1.66,0.92,2.02,'YOGATHY MOCKUP',GREEN);
  head(s,1.66,1.29,4.59,1.08,[['Mockup of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,1.71,4.25,5.52,0.9,L8,GREY);
  kick(s,1.71,3.95,1.39,'MEDITATION',GREEN);
  body(s,1.71,5.53,5.52,0.9,L8,GREY);
  kick(s,1.71,5.22,1.44,'METABOLISM',GREEN);
  kick(s,-2.16,3.69,1.29,'BREATHING',W);
  kick(s,-0.3,3.69,1.39,'MEDITATION',W);
  kick(s,1.65,3.65,1.44,'METABOLISM',W);
}

/** Slide 43 — mockup — laptop */
function slide43(s) {
  // layout 38 furniture
  sh(s,'rect',10,0,3.33,7.5,{fill:GREEN});
  photo(s,4,0.14,11.47,6.93,{flipH:true});
  sh(s,'roundRect',1.32,0.74,4.59,2.04,{fill:mid(LIME,GREEN),shadow:SHADOW,rectRadius:0.29});
  kick(s,1.53,0.97,2.02,'YOGATHY MOCKUP',W);
  head(s,1.53,1.33,4.59,1.08,'Mockup of yogathy help and guide you',null,W);
  sh(s,'roundRect',0.33,4.5,3.29,2.08,{fill:W,shadow:SHADOW,rectRadius:0.35});
  body(s,0.53,4.96,2.87,1.46,L6,GREY);
  kick(s,0.53,4.66,1.29,'BREATHING',GREEN);
  sh(s,'roundRect',3.94,4.5,3.29,2.08,{fill:W,shadow:SHADOW,rectRadius:0.35});
  body(s,4.14,4.96,2.87,1.46,L6,GREY);
  kick(s,4.14,4.66,1.39,'MEDITATION',GREEN);
  body(s,1.46,3.08,4.2,1.18,L8,GREY);
  social(s,11.15,0.16,0.31,LIME);
  button(s,2.49,6.47,1.78,'MORE DETAILS',mid(LIME,GREEN),W,LIME);
  url(s,10.99,6.96,W);
}

/** Slide 44 — mockup — watches */
function slide44(s) {
  // layout 39 furniture
  sh(s,'rect',0,2.73,7.63,1.54,{fill:GREEN});
  photo(s,0.5,0.94,2.84,5.15);
  sh(s,'roundRect',0.71,4.99,2.3,1.13,{fill:MIST,rectRadius:0.19});
  sh(s,'ellipse',0.9,0.9,1.88,0.53,{fill:MIST});
  sh(s,'roundRect',0.67,1.27,2.37,0.81,{fill:MIST,rectRadius:0.14});
  photo(s,3.67,1.57,2.14,3.89);
  sh(s,'roundRect',3.82,4.63,1.74,0.85,{fill:MIST,rectRadius:0.14});
  sh(s,'ellipse',3.96,1.54,1.42,0.4,{fill:MIST});
  sh(s,'roundRect',3.79,1.82,1.78,0.61,{fill:MIST,rectRadius:0.1});
  photo(s,6.11,1.96,1.71,3.1);
  sh(s,'roundRect',6.23,4.4,1.39,0.68,{fill:MIST,rectRadius:0.11});
  sh(s,'ellipse',6.34,1.94,1.13,0.32,{fill:MIST});
  sh(s,'roundRect',6.21,2.16,1.43,0.49,{fill:MIST,rectRadius:0.08});
  sh(s,'roundRect',8.06,1.27,4.59,2.04,{fill:mid(LIME,GREEN),shadow:SHADOW,rectRadius:0.29});
  kick(s,8.26,1.5,2.02,'YOGATHY MOCKUP',W);
  head(s,8.26,1.86,4.59,1.08,'Mockup of yogathy help and guide you',null,W);
  body(s,8.27,3.75,4.19,1.46,L10,GREY);
  url(s,1.02,6.96,AMBER);
  social(s,11.9,0.21,0.31,GREEN);
  stamp(s,9.68,6.93,AMBER);
  logo(s,0.3,0.31,GREEN);
}

/** Slide 45 — mockup — phones */
function slide45(s) {
  // layout 40 furniture
  sh(s,'rect',8.89,0,3.33,7.5,{fill:GREEN});
  photo(s,8.17,3.82,1.72,3.19);
  photo(s,8.16,0.54,1.72,3.19);
  photo(s,9.99,0.98,3,5.59);
  kick(s,1.49,1.59,2.02,'YOGATHY MOCKUP',GREEN);
  head(s,1.49,1.96,4.59,1.08,[['Mockup of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,1.49,3.23,4.86,1.74,L1,GREY);
  url(s,0.43,6.89,AMBER);
  line(s,0,0.91,2.25,AMBER,6);
  button(s,3.97,5.19,1.78,'MORE DETAILS',mid(LIME,GREEN),W,LIME);
  social(s,10.08,0.16,0.31,LIME);
  stamp(s,8.93,6.96,W);
  logo(s,0.3,0.31,GREEN);
}

/** Slide 46 — infographic — thermometers */
function slide46(s) {
  sh(s,'ellipse',11.05,2.87,1.07,3.95,{fill:{color:CLOUD,transparency:35}});
  sh(s,'ellipse',11.29,2.75,0.58,0.23,{fill:SILVER});
  sh(s,'ellipse',9.84,2.87,1.07,3.95,{fill:{color:CLOUD,transparency:35}});
  sh(s,'ellipse',10.08,2.75,0.58,0.23,{fill:SILVER});
  sh(s,'ellipse',8.62,2.87,1.07,3.95,{fill:{color:CLOUD,transparency:35}});
  sh(s,'ellipse',8.87,2.75,0.58,0.23,{fill:SILVER});
  sh(s,'ellipse',7.43,2.87,1.07,3.95,{fill:{color:CLOUD,transparency:35}});
  sh(s,'ellipse',7.67,2.75,0.58,0.23,{fill:SILVER});
  sh(s,'ellipse',7.49,5.72,0.95,1.02,{fill:mid(GOLD,AMBER)});
  sh(s,'custGeom',7.73,3.6,0.47,2.38,{fill:GOLD,points:[{x:0,y:0},{x:0.47,y:0},{x:0.47,y:2.38},{x:0,y:2.38},{x:0,y:0}]});
  sh(s,'ellipse',7.73,3.49,0.47,0.23,{fill:SAND});
  sh(s,'ellipse',8.69,5.72,0.95,1.02,{fill:mid(LIME,GREEN)});
  sh(s,'custGeom',8.93,3.94,0.47,2.04,{fill:LIME,points:[{x:0,y:0},{x:0.47,y:0},{x:0.47,y:2.04},{x:0,y:2.04},{x:0,y:0}]});
  sh(s,'ellipse',8.93,3.86,0.47,0.23,{fill:MINT});
  sh(s,'ellipse',9.9,5.72,0.95,1.02,{fill:mid(AMBER,BRONZE)});
  sh(s,'custGeom',10.14,3.6,0.47,2.38,{fill:AMBER,points:[{x:0,y:0},{x:0.47,y:0},{x:0.47,y:2.38},{x:0,y:2.38},{x:0,y:0}]});
  sh(s,'ellipse',10.14,3.5,0.47,0.23,{fill:GOLD});
  sh(s,'ellipse',11.11,5.72,0.95,1.02,{fill:mid(GREEN,FOREST)});
  sh(s,'custGeom',11.35,4.12,0.47,1.87,{fill:GREEN,points:[{x:0,y:0},{x:0.47,y:0},{x:0.47,y:1.87},{x:0,y:1.87},{x:0,y:0}]});
  sh(s,'ellipse',11.35,4,0.47,0.23,{fill:LIME});
  txt(s,7.68,6.08,0.6,0.35,'94%',{color:W,face:PP,size:15,align:'center',wrap:false});
  txt(s,8.89,6.08,0.58,0.35,'78%',{color:W,face:PP,size:15,align:'center',wrap:false});
  txt(s,10.09,6.08,0.6,0.35,'88%',{color:W,face:PP,size:15,align:'center',wrap:false});
  txt(s,11.33,6.08,0.6,0.35,'68%',{color:W,face:PP,size:15,align:'center',wrap:false});
  kick(s,1.11,0.79,2.5,'YOGATHY INFOGRAPHIC',GREEN);
  head(s,1.11,1.16,4.92,0.59,[['Infographic of '],['yogathy',AMBER]]);
  sh(s,'ellipse',1.15,3.57,0.73,0.73,{fill:MINT});
  sh(s,'ellipse',1.18,3.6,0.66,0.66,{fill:LIME});
  sh(s,'ellipse',1.22,3.64,0.59,0.59,{fill:GREEN});
  sh(s,'ellipse',1.34,3.76,0.34,0.37,{fill:W});
  sh(s,'ellipse',1.15,5.92,0.73,0.73,{fill:MINT});
  sh(s,'ellipse',1.18,5.95,0.66,0.66,{fill:LIME});
  sh(s,'ellipse',1.22,5.99,0.59,0.59,{fill:GREEN});
  sh(s,'ellipse',1.4,6.05,0.22,0.47,{fill:W,rotate:45});
  sh(s,'ellipse',1.15,2.46,0.73,0.73,{fill:MINT});
  sh(s,'ellipse',1.18,2.5,0.66,0.66,{fill:LIME});
  sh(s,'ellipse',1.22,2.54,0.59,0.59,{fill:GREEN});
  sh(s,'ellipse',1.32,2.64,0.38,0.37,{fill:W});
  sh(s,'ellipse',1.18,4.77,0.73,0.73,{fill:MINT});
  sh(s,'ellipse',1.21,4.81,0.66,0.66,{fill:LIME});
  sh(s,'ellipse',1.25,4.84,0.59,0.59,{fill:GREEN});
  sh(s,'ellipse',1.4,4.88,0.28,0.51,{fill:W,rotate:45});
  body(s,1.96,2.55,4.49,0.62,L7,GREY);
  kick(s,1.96,2.25,0.72,'FIRST',GREEN);
  body(s,1.96,3.78,4.49,0.62,L7,GREY);
  kick(s,1.96,3.48,0.99,'SECOND',GREEN);
  body(s,1.96,5,4.49,0.62,L7,GREY);
  kick(s,1.96,4.7,0.77,'THIRD',GREEN);
  body(s,1.96,6.23,4.49,0.62,L7,GREY);
  kick(s,1.96,5.92,0.96,'FOURTH',GREEN);
  social(s,11.9,0.21,0.31,GREEN);
  stamp(s,9.68,6.93,AMBER);
}

/** Slide 47 — infographic — petals */
function slide47(s) {
  sh(s,'teardrop',9.85,2.36,2.18,2.18,{fill:MIST,rotate:180});
  sh(s,'teardrop',9.97,2.48,1.93,1.93,{fill:FOREST,rotate:180});
  sh(s,'teardrop',9.85,4.58,2,2,{fill:MIST,rotate:180});
  sh(s,'teardrop',9.96,4.7,1.78,1.78,{fill:AMBER,rotate:180});
  sh(s,'teardrop',8.4,3.14,1.39,1.39,{fill:MIST,rotate:180});
  sh(s,'teardrop',8.48,3.22,1.24,1.24,{fill:GOLD,rotate:180});
  sh(s,'teardrop',8.14,4.58,1.65,1.65,{fill:MIST,rotate:180});
  sh(s,'teardrop',8.24,4.68,1.47,1.47,{fill:LIME,rotate:180});
  kick(s,1.43,0.88,2.5,'YOGATHY INFOGRAPHIC',GREEN);
  head(s,1.43,1.25,4.92,0.59,[['Infographic of '],['yogathy',AMBER]]);
  sh(s,'ellipse',8.67,5.07,0.59,0.63,{fill:W});
  sh(s,'ellipse',10.58,5.18,0.45,0.94,{fill:W,rotate:45});
  sh(s,'ellipse',10.4,3.06,0.9,0.88,{fill:W});
  sh(s,'ellipse',8.94,3.52,0.35,0.64,{fill:W,rotate:45});
  body(s,1.43,2.43,4.49,0.62,L7,GREY);
  kick(s,1.43,2.13,0.72,'FIRST',GREEN);
  body(s,1.43,3.66,4.49,0.62,L7,GREY);
  kick(s,1.43,3.36,0.99,'SECOND',GREEN);
  body(s,1.43,4.9,4.49,0.62,L7,GREY);
  kick(s,1.43,4.59,0.77,'THIRD',GREEN);
  body(s,1.43,6.13,4.49,0.62,L7,GREY);
  kick(s,1.43,5.83,0.96,'FOURTH',GREEN);
  url(s,0.43,6.89,AMBER);
  logo(s,0.3,0.31,GREEN);
  stamp(s,9.68,6.93,AMBER);
  line(s,11.1,0.91,2.25,AMBER,6);
}

/** Slide 48 — SWOT — strength */
function slide48(s) {
  // layout 43 furniture
  sh(s,'rect',10,0,3.33,7.5,{fill:GREEN});
  line(s,4.87,6.08,2.25,AMBER,6);
  kick(s,0.95,2.1,2.17,'YOGATHY STRENGTH',GREEN);
  head(s,0.95,2.47,4.59,1.08,[['Strength of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,0.95,3.47,4.86,1.74,L1,GREY);
  url(s,0.43,6.89,AMBER);
  logo(s,0.3,0.31,GREEN);
  stamp(s,9.15,6.05,W);
}

/** Slide 49 — SWOT — weakness */
function slide49(s) {
  // layout 44 furniture
  sh(s,'rect',7.11,0.89,5.6,5.71,{fill:GREEN,shadow:{type:'outer',blur:15,offset:0,angle:0,color:BLACK,opacity:0.18}});
  kick(s,7.46,1.69,2.22,'YOGATHY WEAKNESS',W);
  head(s,7.46,2.05,4.59,1.08,'Weakness of yogathy help and guide you',null,W);
  body(s,7.46,3.06,4.86,1.74,L1,W);
  button(s,9.96,5.82,1.78,'READ MORE',W,GREEN,LIME);
  social(s,0.64,0.27,0.31,GREEN);
  url(s,0.43,6.89,AMBER);
  stamp(s,9.68,6.93,AMBER);
}

/** Slide 50 — SWOT — opportunity */
function slide50(s) {
  sh(s,'roundRect',0.92,0.89,4.59,2.04,{fill:mid(LIME,GREEN),shadow:SHADOW,rectRadius:0.29});
  kick(s,1.12,1.12,2.52,'YOGATHY OPPORTUNITY',W);
  head(s,1.12,1.49,3.96,1.08,'Opportunity of yogathy help you',null,W);
  sh(s,'roundRect',0.92,3.16,4.46,3.9,{fill:W,shadow:SHADOW,rectRadius:0.65});
  body(s,1.18,3.74,3.89,1.46,L6,GREY);
  kick(s,1.18,3.44,1.75,'BREATHING',GREEN);
  body(s,1.18,5.31,3.89,1.46,L6,GREY);
  kick(s,1.18,5.01,1.39,'MEDITATION',GREEN);
  button(s,3.42,6.42,1.78,'MORE DETAILS',mid(LIME,GREEN),W,LIME);
  social(s,11.78,0.27,0.31,GREEN);
  url(s,11.05,6.89,AMBER);
  logo(s,0.3,0.31,GREEN);
}

/** Slide 51 — SWOT — threats */
function slide51(s) {
  // layout 46 furniture
  sh(s,'rect',1.37,0,3.33,7.5,{fill:GREEN});
  line(s,6.67,1.72,2.25,AMBER,6);
  kick(s,7.59,2.2,2.02,'YOGATHY THREATS',GREEN);
  head(s,7.59,2.57,4.59,1.08,[['Threats of '],['yogathy',AMBER],[' help and guide you']]);
  body(s,7.59,3.62,4.86,1.74,L1,GREY);
  stamp(s,1.25,1.38,W);
  url(s,10.7,6.89,AMBER);
}

/** Slide 52 — pricing — three green plans */
function slide52(s) {
  sh(s,'roundRect',0.58,1.46,3.65,5.7,{fill:mid(MINT,GREEN),shadow:SHADOW,rectRadius:0.51});
  sh(s,'roundRect',4.84,1.46,3.65,5.7,{fill:mid(MINT,GREEN),shadow:SHADOW,rectRadius:0.51});
  sh(s,'roundRect',9.11,1.46,3.65,5.7,{fill:mid(MINT,GREEN),shadow:SHADOW,rectRadius:0.51});
  kick(s,5.11,0.37,2.62,'YOGATHY PRICING PLANS',W,'center');
  head(s,3.52,0.74,5.79,0.59,'Pricing plans of yogathy','center',W);
  kick(s,1.71,3.92,1.39,'BASIC YOGA',W,'center');
  num44(s,1.48,2.52,1.99,0.84,'165$',W,'center');
  sh(s,'roundRect',1.51,1.76,1.78,0.49,{fill:AMBER,rectRadius:0.24});
  chip(s,1.88,1.8,1.04,'BASIC',W,'center');
  button(s,1.22,6.53,1.78,'CHOOSE NOW',W,GREEN,LIME);
  sh(s,'roundRect',5.76,1.76,1.82,0.49,{fill:AMBER,rectRadius:0.24});
  chip(s,5.82,1.8,1.7,'STANDARD',W,'center');
  sh(s,'roundRect',10.04,1.76,1.78,0.49,{fill:AMBER,rectRadius:0.24});
  chip(s,10.55,1.8,0.76,'PRO',W,'center');
  social(s,11.78,0.27,0.31,GREEN);
  logo(s,0.3,0.31,GREEN);
  line(s,1.27,4.32,2.25,W,2.25);
  num44(s,5.75,2.52,1.99,0.84,'185$',W,'center');
  button(s,5.49,6.53,1.78,'CHOOSE NOW',W,GREEN,LIME);
  num44(s,10.07,2.52,1.99,0.84,'200$',W,'center');
  button(s,9.81,6.53,1.78,'CHOOSE NOW',W,GREEN,LIME);
  body(s,0.76,4.54,3.21,1.18,L3,W,'center');
  kick(s,5.75,3.92,1.84,'STANDARD YOGA',W,'center');
  line(s,5.55,4.32,2.25,W,2.25);
  body(s,5.03,4.54,3.21,1.18,L3,W,'center');
  kick(s,10.37,3.92,1.19,'PRO YOGA',W,'center');
  line(s,9.83,4.32,2.25,W,2.25);
  body(s,9.31,4.54,3.21,1.18,L3,W,'center');
}

/** Slide 53 — pricing — three white plans */
function slide53(s) {
  sh(s,'roundRect',1.83,1.43,3.65,5.7,{fill:W,shadow:SHADOW,rectRadius:0.51});
  sh(s,'roundRect',8.09,1.43,3.65,5.7,{fill:W,shadow:SHADOW,rectRadius:0.51});
  sh(s,'roundRect',1.83,1.43,3.65,5.7,{fill:{color:LIME,transparency:74},rectRadius:0.5});
  sh(s,'roundRect',8.09,1.43,3.65,5.7,{fill:{color:LIME,transparency:74},rectRadius:0.5});
  kick(s,5.11,0.2,2.62,'YOGATHY PRICING PLANS',W,'center');
  head(s,3.52,0.57,5.79,0.59,'Pricing plans of yogathy','center',W);
  sh(s,'roundRect',2.63,6.49,1.78,0.49,{fill:AMBER,rectRadius:0.24});
  chip(s,2.85,6.53,1.35,'CHOOSE',W,'center');
  sh(s,'roundRect',9.02,6.49,1.78,0.49,{fill:AMBER,rectRadius:0.24});
  chip(s,9.24,6.53,1.35,'CHOOSE',W,'center');
  kick14(s,2.14,2.29,1.76,[['BASIC'],[' YOGA ',GREY,{breakLine:true}],['& MEDITATION',GREY]],GREEN);
  line(s,2.16,2.99,2.25,AMBER,2.25);
  body(s,2.04,3.47,2.92,1.46,L3,GREY);
  num44(s,1.89,5.32,1.99,0.84,'165$',GREEN,'center');
  kick14(s,8.63,2.29,1.76,[['PRO'],[' YOGA ',GREY,{breakLine:true}],['& MEDITATION',GREY]],GREEN);
  line(s,8.65,2.99,2.25,AMBER,2.25);
  body(s,8.53,3.47,2.92,1.46,L3,GREY);
  num44(s,8.37,5.32,1.99,0.84,'200$',GREEN,'center');
  kick14(s,3.55,5.65,0.81,'/ DAY',GREY);
  kick14(s,10.09,5.65,0.81,'/ DAY',GREY);
  social(s,11.78,0.27,0.31,GREEN);
  logo(s,2.18,1.7,GREEN);
  logo(s,8.49,1.7,GREEN);
  sh(s,'roundRect',4.96,1.22,3.65,6.1,{fill:W,shadow:SHADOW,rectRadius:0.51});
  sh(s,'roundRect',5.9,6.49,1.78,0.49,{fill:AMBER,rectRadius:0.24});
  chip(s,6.11,6.53,1.35,'CHOOSE',W,'center');
  txt(s,5.31,1.88,2.38,0.64,[['STANDARD'],[' YOGA ',GREY,{breakLine:true}],['& MEDITATION',GREY]],{color:GREEN,face:PP,size:16,bold:true,spc:1,wrap:false});
  line(s,5.33,2.65,2.25,AMBER,2.25);
  body(s,5.21,3.24,2.92,1.46,L3,GREY);
  num44(s,5.05,5.32,1.99,0.84,'185$',GREEN,'center');
  kick14(s,6.77,5.65,0.81,'/ DAY',GREY);
  logo(s,5.31,1.38,GREEN);
}

/** Slide 54 — pricing — two plans */
function slide54(s) {
  sh(s,'roundRect',0.71,2.71,3.96,2.83,{fill:mid(MINT,GREEN),shadow:SHADOW,rectRadius:0.4});
  sh(s,'roundRect',5,2.71,3.96,2.83,{fill:W,shadow:SHADOW,rectRadius:0.4});
  kick(s,3.21,1.13,2.62,'YOGATHY PRICING PLANS',AMBER,'center');
  head(s,1.62,1.5,5.79,0.59,[['Pricing plans of '],['yogathy',AMBER]],'center',BLACK);
  kick(s,2.02,3.46,1.36,'DAILY YOGA',W,'center');
  body(s,1.06,4.04,3.21,1.18,L3,W,'center');
  line(s,1.51,3.92,2.25,W,2.25);
  num44(s,1.69,2.71,1.99,0.84,'165$',W,'center');
  kick(s,6.24,3.46,1.45,'WEEKS YOGA',GREEN,'center');
  body(s,5.33,4.04,3.21,1.18,L3,GREY,'center');
  line(s,5.78,3.92,2.25,GREEN,2.25);
  num44(s,5.96,2.71,1.99,0.84,'200$',BLACK,'center');
  sh(s,'ellipse',3.68,5.26,0.89,0.79,{fill:mid(MINT,GREEN),shadow:SHADOW});
  kick14(s,3.8,5.4,0.64,'BUY',W,'center');
  arrowBadge(s,4.03,5.68,0.2,W,LIME);
  sh(s,'ellipse',5.07,5.26,0.89,0.79,{fill:mid(MINT,GREEN),shadow:SHADOW});
  kick14(s,5.19,5.4,0.64,'BUY',W,'center');
  arrowBadge(s,5.42,5.68,0.2,W,LIME);
  logo(s,0.3,0.31,GREEN);
  url(s,0.43,6.89,AMBER);
}

/** Slide 55 — testimonials */
function slide55(s) {
  // layout 48 furniture
  sh(s,'roundRect',0.39,2.56,3.96,3.22,{fill:mid(MINT,GREEN),shadow:SHADOW,rectRadius:0.45});
  sh(s,'roundRect',4.69,2.56,3.96,3.22,{fill:W,shadow:SHADOW,rectRadius:0.45});
  sh(s,'roundRect',8.98,2.56,3.96,3.22,{fill:mid(MINT,GREEN),shadow:SHADOW,rectRadius:0.45});
  kick(s,5.28,0.83,2.44,'YOGATHY TESTIMONIAL',GREEN,'center');
  head(s,4.38,1.2,4.59,1.08,[['What clients say about '],['yogathy',AMBER]],'center');
  body(s,0.76,4.16,3.21,1.18,L3,W,'center');
  body(s,9.37,4.16,3.21,1.18,L3,W,'center');
  body(s,5.07,4.16,3.21,1.18,L3,GREY,'center');
  social(s,6.12,3.85,0.31,GREEN);
  social(s,1.81,3.84,0.31,LIME);
  social(s,10.41,3.84,0.31,LIME);
  logo(s,0.3,0.31,GREEN);
  url(s,10.7,6.89,AMBER);
  arrowBadge(s,8.49,3.95,0.31,W,LIME);
  arrowBadge(s,4.56,3.95,0.31,W,LIME);
}

/** Slide 56 — schedule — calendar chips */
function slide56(s) {
  // layout 49 furniture
  sh(s,'rect',0,0,2.23,7.51,{fill:GREEN});
  sh(s,'rect',4.93,0.42,1.67,1.54,{fill:GREEN});
  sh(s,'rect',5.07,0.98,1.38,0.89,{fill:W});
  num44(s,4.76,1.07,1.99,0.84,'01',GREEN,'center');
  kick14(s,5.35,0.57,0.81,'APRIL',W,'center');
  sh(s,'roundRect',4.4,0.52,0.67,0.23,{fill:GREEN,line:{color:W,width:3},rectRadius:0.11});
  sh(s,'rect',4.93,2.98,1.67,1.54,{fill:GREEN});
  sh(s,'rect',5.07,3.54,1.38,0.89,{fill:W});
  num44(s,4.76,3.63,1.99,0.84,'21',GREEN,'center');
  kick14(s,5.35,3.14,0.81,'APRIL',W,'center');
  sh(s,'roundRect',4.4,3.08,0.67,0.23,{fill:GREEN,line:{color:W,width:3},rectRadius:0.11});
  sh(s,'rect',4.93,5.45,1.67,1.54,{fill:GREEN});
  sh(s,'rect',5.07,6.01,1.38,0.89,{fill:W});
  num44(s,4.76,6.1,1.99,0.84,'28',GREEN,'center');
  kick14(s,5.35,5.61,0.81,'APRIL',W,'center');
  sh(s,'roundRect',4.4,5.55,0.67,0.23,{fill:GREEN,line:{color:W,width:3},rectRadius:0.11});
  body(s,7.18,0.74,3.89,1.46,L6,GREY);
  kick(s,7.18,0.44,1.58,'REGISTRATION',GREEN);
  body(s,7.18,3.14,3.89,1.46,L6,GREY);
  kick(s,7.18,2.84,1.65,'CHOOSE CLASS',GREEN);
  body(s,7.18,5.53,3.89,1.46,L6,GREY);
  kick(s,7.18,5.23,1.4,'START YOGA',GREEN);
  social(s,11.78,0.27,0.31,GREEN);
  url(s,10.7,6.89,AMBER);
}

/** Slide 57 — contact — wide card */
function slide57(s) {
  sh(s,'roundRect',1.59,4.01,10.15,2.35,{fill:mid(LIME,GREEN),shadow:SHADOW,rectRadius:0.33});
  small(s,5.85,4.6,1.85,0.63,[['619 Queens, LA ',null,{breakLine:true}],['US 555 ']],MIST);
  kick14(s,5.87,4.38,1.75,'LOCATION',W);
  small(s,9.07,4.52,1.76,0.63,[['Monday – Friday',null,{breakLine:true}],['10.00 AM – 03.00 PM']],MIST);
  kick14(s,9.08,4.29,1.75,'OFFICE HOURS',W);
  small(s,9.08,5.69,1.75,0.35,'@yogathy.com',MIST);
  kick14(s,9.09,5.47,1.75,'EMAIL',W);
  txt(s,5.85,5.74,1.23,0.35,'+384 829 7262',{color:MIST,face:NS,size:11,lsp:1.5});
  kick14(s,5.87,5.52,1.75,'PHONE',W);
  kick(s,4.96,2.39,2.5,'YOGATHY INFOGRAPHIC',W);
  head(s,4.96,2.75,4.25,0.59,'Contact of yogathy',null,W);
  arrowBadge(s,10.92,5.64,0.58,W,GREEN);
  url(s,0.43,6.89,AMBER);
  stamp(s,9.68,6.93,AMBER);
}

/** Slide 58 — contact — left card */
function slide58(s) {
  sh(s,'roundRect',0.83,0.72,5.19,5.58,{fill:mid(LIME,GREEN),shadow:SHADOW,rectRadius:0.72});
  kick(s,1.33,1.26,2.5,'YOGATHY INFOGRAPHIC',W);
  head(s,1.33,1.63,4.25,0.59,'Contact of yogathy',null,W);
  small(s,1.33,3.53,1.85,0.63,[['619 Queens, LA ',null,{breakLine:true}],['US 555 ']],MIST);
  kick14(s,1.34,3.31,1.75,'LOCATION',W);
  small(s,3.82,3.44,1.76,0.63,[['Monday – Friday',null,{breakLine:true}],['10.00 AM – 03.00 PM']],MIST);
  kick14(s,3.83,3.22,1.75,'OFFICE HOURS',W);
  small(s,3.83,4.62,1.75,0.35,'@yogathy.com',MIST);
  kick14(s,3.85,4.4,1.75,'EMAIL',W);
  txt(s,1.33,4.67,1.23,0.35,'+384 829 7262',{color:MIST,face:NS,size:11,lsp:1.5});
  kick14(s,1.34,4.45,1.75,'PHONE',W);
  small(s,1.33,2.3,3.24,0.63,[['Please feel free to call us on (phone cell) or contact us by (email),']],MIST,'justify');
  arrowBadge(s,5.43,5.79,0.58,W,GREEN);
  social(s,2.89,5.8,0.31,GREEN);
  url(s,0.43,6.89,W);
  stamp(s,9.68,6.93,W);
}

/** Slide 59 — thanks — light */
function slide59(s) {
  hero(s,0.76,2.02,5.83,1.58,'THANKS',BLACK,'center');
  tag(s,0.89,3.42,3.76,'KEEP YOUR BODY HEALTHY',AMBER);
  line(s,0,5.17,2.25,GREEN,6);
  url(s,11.23,0.18,AMBER);
  social(s,0.33,0.48,0.43,GREEN);
  stamp(s,9.68,6.93,W);
}

/** Slide 60 — thanks — green wash */
function slide60(s) {
  fade(s,0,0.02,13.33,7.5,LIME,0.27,0.52);
  hero(s,6.51,1.84,5.83,1.58,'THANKS',W,'center');
  tag(s,9.14,6.76,3.76,'KEEP YOUR BODY HEALTHY',W);
  url(s,0.43,6.95,W);
  social(s,11.34,0.48,0.43,GREEN);
  line(s,0,1.46,2.25,W,6);
  body(s,6.8,3.16,4.03,1.18,L14,W);
}

/** Slide 61 — thanks — right photo */
function slide61(s) {
  hero(s,7.39,2.31,5.83,1.58,'THANKS',GREEN,'center');
  logo(s,0.3,0.31,GREEN);
  stamp(s,9.54,6.88,AMBER);
  url(s,11.08,0.18,AMBER);
  body(s,8.1,4.06,4.41,0.9,L17,GREY,'center');
  line(s,9.18,3.93,2.25,AMBER,6);
}

/** Slide 62 — thanks — green column */
function slide62(s) {
  sh(s,'rect',8.72,0,3,7.5,{fill:{color:GREEN,transparency:32}});
  hero(s,4.68,3.24,5.83,1.58,'THANKS',W,'right');
  tag(s,6.55,4.82,3.76,'KEEP YOUR BODY HEALTHY',W);
  url(s,9.63,6.95,W);
  stamp(s,0.58,6.96,W);
  social(s,0.44,0.48,0.43,GREEN);
}

/* ------------------------------------------------------------------- build */
const BUILDERS = [
  slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8,
  slide9, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
  slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
  slide25, slide26, slide27, slide28, slide29, slide30, slide31, slide32,
  slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40,
  slide41, slide42, slide43, slide44, slide45, slide46, slide47, slide48,
  slide49, slide50, slide51, slide52, slide53, slide54, slide55, slide56,
  slide57, slide58, slide59, slide60, slide61, slide62,
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'YOGATHY', width: 13.333, height: 7.5 });
pptx.layout = 'YOGATHY';
pptx.author = 'Yoga Healthy Corp.';
pptx.title = 'Yogathy';

BUILDERS.forEach((build) => {
  const slide = pptx.addSlide();
  slide.background = { color: W };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '098c4285-8e5b-46f0-9d7c-eb0efe97f6c2_grok_final.pptx') })
  .then((f) => console.log('wrote ' + f));
