/**
 * AiTech — "How AI Technology is Transforming the World" (60-slide template deck)
 * Standalone pptxgenjs recreation.  Run:  node <thisfile>.js
 *
 * Design system
 *   canvas   13.333 x 7.5 in (16:9)
 *   headings Poppins Medium   body Nunito
 *   ink #262626 · grey #808080 · indigo #4437E6 · teal #00C7AA · white #FFFFFF
 *
 * Photographs in the source deck are replaced by flat grey placeholder cards.
 */
'use strict';
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ---------------------------------------------------------------- palette */
const INK = '262626';       // near-black headline / body ink
const INK2 = '404040';
const GREY = '808080';      // secondary body copy
const W = 'FFFFFF';
const INDIGO = '4437E6';    // theme accent1
const TEAL = '00C7AA';      // theme accent2
const HAIR = 'F2F2F2';      // hairline decoration rings
const PHOTO = 'ECEEF3';     // picture-slot outline
const PHOTO_LBL = 'C6CAD4'; // picture-slot caption

/* ---------------------------------------------------------------- fonts */
const HEAD = 'Poppins Medium';
const BODY = 'Nunito';

/* PowerPoint's default text-box insets: 0.1" sides, 0.05" top/bottom (in points) */
const PAD = [7.2, 7.2, 3.6, 3.6];

/* soft ambient shadow used by every raised card in the deck
   (a fresh object each call - pptxgenjs rewrites the one it is given) */
function shadow() {
  return { type: 'outer', color: '000000', opacity: 0.1, blur: 15, offset: 3, angle: 90 };
}

/* ---------------------------------------------------------------- helpers */

// rectangle / rounded card:  card(slide, x, y, w, h, fillColor, cornerRadius, extras)
// extras — {line, lw, shadow, top} where `top` rounds only the two top corners
function card(s, x, y, w, h, fill, r, o) {
  o = o || {};
  s.addShape(o.top ? 'round2SameRect' : (r ? 'roundRect' : 'rect'), {
    x: x, y: y, w: w, h: h,
    rectRadius: r || undefined,
    fill: fill ? { color: fill } : { type: 'none' },
    line: o.line ? { color: o.line, width: o.lw || 1 } : { type: 'none' },
    shadow: o.shadow ? shadow() : undefined,
    flipV: o.top ? true : undefined,
  });
}

function oval(s, x, y, w, h, fill, o) {
  o = o || {};
  s.addShape('ellipse', {
    x: x, y: y, w: w, h: h,
    fill: fill ? { color: fill } : { type: 'none' },
    line: o.line ? { color: o.line, width: o.lw || 1 } : { type: 'none' },
    shadow: o.shadow ? shadow() : undefined,
  });
}

function hline(s, x, y, w, color) {
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: color, width: 0.75 } });
}

/* text runs: 'string' or [[text, color, {sz,f,b,i,br}], ...] */
function runs(spec, o) {
  const list = Array.isArray(spec) ? spec : [[spec, o.c || INK]];
  return list.map(function (r) {
    const e = r[2] || {};
    return {
      text: r[0],
      options: {
        fontSize: e.sz || o.sz || 14,
        fontFace: e.f || o.f || HEAD,
        color: r[1],
        bold: !!e.b, italic: !!e.i,
        charSpacing: e.sp || o.sp || undefined,
        breakLine: !!e.br,
      },
    };
  });
}

function txt(s, x, y, w, h, spec, o) {
  o = o || {};
  s.addText(runs(spec, o), {
    x: x, y: y, w: w, h: h,
    align: o.al === 'c' ? 'center' : 'left',
    valign: o.va === 'm' ? 'middle' : 'top',
    lineSpacing: o.lp || undefined,
    lineSpacingMultiple: o.ls || undefined,
    margin: PAD, wrap: true, isTextBox: true,
  });
}

/* 12 pt Nunito paragraph at 150 % leading — the deck's default body copy */
function body(s, x, y, w, h, spec, color, o) {
  o = Object.assign({ sz: 12, f: BODY, ls: 1.5, c: color || GREY }, o || {});
  txt(s, x, y, w, h, spec, o);
}

/* filled/outlined shape that carries its own centred label */
function pill(s, x, y, w, h, spec, o) {
  o = o || {};
  s.addText(runs(spec, o), {
    x: x, y: y, w: w, h: h,
    shape: o.r ? 'roundRect' : 'rect',
    rectRadius: o.r || undefined,
    fill: o.fill ? { color: o.fill } : { type: 'none' },
    line: o.line ? { color: o.line, width: 1 } : { type: 'none' },
    shadow: o.shadow ? shadow() : undefined,
    align: o.al === 'c' ? 'center' : 'left',
    valign: o.va === 'm' ? 'middle' : 'top',
    lineSpacing: o.lp || undefined,
    lineSpacingMultiple: o.ls || undefined,
    margin: PAD, wrap: true,
  });
}

/* pictogram stand-in for the deck's line-art glyphs: an outlined rounded square
   with a small solid dot, tinted with the glyph's own colour */
function icon(s, x, y, w, h, color) {
  const c = color || INK;
  const d = Math.min(w, h) * 0.82;
  const gx = x + (w - d) / 2, gy = y + (h - d) / 2;
  s.addShape('roundRect', { x: gx, y: gy, w: d, h: d, rectRadius: d * 0.25,
    fill: { type: 'none' }, line: { color: c, width: 1.25 } });
  s.addShape('ellipse', { x: gx + d * 0.34, y: gy + d * 0.34, w: d * 0.32, h: d * 0.32,
    fill: { color: c }, line: { type: 'none' } });
}

/* picture slot: outlined frame + "[image]" caption, standing in for a photograph */
function photo(s, x, y, w, h, r) {
  const rad = r === undefined ? Math.min(w, h) * 0.05 : r;
  s.addShape(rad > 0 ? 'roundRect' : 'rect', {
    x: x, y: y, w: w, h: h,
    rectRadius: rad || undefined,
    fill: { type: 'none' }, line: { color: PHOTO, width: 1 },
  });
  if (w >= 1.4 && h >= 0.9) {
    s.addText('[image]', {
      x: x, y: y, w: w, h: h, align: 'center', valign: 'middle',
      fontSize: 10, fontFace: BODY, color: PHOTO_LBL, margin: 0,
    });
  }
}

/* Device mock-up for the "Device Mockup" slides, drawn from primitives.
   Every number is a fraction of the handset's own width (u) or length (span),
   measured off the original artwork. */
function phone(s, x, y, w, h, landscape) {
  // the artwork sits slightly inside its bounding box on the short axis
  if (landscape) { y += h * 0.008; h *= 0.9835; } else { x += w * 0.008; w *= 0.9835; }
  const u = landscape ? h : w;
  const shell = '414041', edge = '131313', btn = '919191';
  const box = function (dx, dy, bw, bh, r, color) {
    s.addShape('roundRect', { x: x + dx, y: y + dy, w: bw, h: bh,
      rectRadius: r, fill: { color: color }, line: { type: 'none' } });
  };
  box(0, 0, w, h, u * 0.148, shell);                                             // outer shell
  box(u * 0.016, u * 0.016, w - u * 0.032, h - u * 0.032, u * 0.134, edge);      // glass edge
  box(u * 0.062, u * 0.062, w - u * 0.124, h - u * 0.124, u * 0.105, 'FFFFFF');  // screen

  // notch hanging from the top (or left) bezel, with speaker slit and camera dot
  const nl = u * 0.494, nTop = u * 0.016, nd = u * 0.134 - nTop;
  const slitL = u * 0.116, slitD = u * 0.017, slitAt = u * 0.072;
  const dot = u * 0.025, dotAt = u * 0.068, dotOff = u * 0.105;
  if (landscape) {
    box(nTop, (h - nl) / 2, nd, nl, u * 0.05, edge);
    box(slitAt, (h - slitL) / 2, slitD, slitL, slitD / 2, shell);
    box(dotAt, (h + dotOff * 2 - dot) / 2, dot, dot, dot / 2, shell);
  } else {
    box((w - nl) / 2, nTop, nl, nd, u * 0.05, edge);
    box((w - slitL) / 2, slitAt, slitL, slitD, slitD / 2, shell);
    box((w + dotOff * 2 - dot) / 2, dotAt, dot, dot, dot / 2, shell);
  }

  // side buttons: silent switch + volume pair on one flank, power on the other
  const t = u * 0.016, span = landscape ? w : h;
  const flank = function (from, to, far) {
    const a = span * from, b = span * to;
    if (landscape) {
      s.addShape('rect', { x: x + a, y: y + (far ? h - t / 2 : -t / 2), w: b - a, h: t,
        fill: { color: btn }, line: { type: 'none' } });
    } else {
      s.addShape('rect', { x: x + (far ? w - t / 2 : -t / 2), y: y + a, w: t, h: b - a,
        fill: { color: btn }, line: { type: 'none' } });
    }
  };
  flank(0.137, 0.175, false);
  flank(0.209, 0.280, false);
  flank(0.299, 0.371, false);
  flank(0.234, 0.345, true);
}
/* ---------------------------------------------------------------- body copy reused across slides */
const L1 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod';
const L2 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim';
const L3 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.';
const L4 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad';
const L5 = 'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor';
const L6 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore';
const L7 = 'Lorem ipsum dolor sit amet, consectetur adipiscing';
const L8 = 'Lorem ipsum dolor sit amet, elit, sed do eiusmod';
const L9 = 'Lorem ipsum dolor amet, consectetur adipiscing elit,.';
const L10 = 'Building a Connection.';
const L11 = 'Lorem ipsum dolor sit amet, consectetur';
const L12 = 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad.';
const L13 = 'Lorem ipsum dolor sit elit, sed do eiusmod tempor ut labore et dolore ';
const L14 = 'The AI Technology Presentation';
const L15 = 'Automation Intelligence Evolution';
const L16 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation.';
const L17 = 'Lorem ipsum dolor sit';
const L18 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et';
const L19 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ut do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis';
const L20 = 'Lorem ipsum dolor amet, consectetur adipiscing ut do eiusmod tempor';
const L21 = 'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor ut et dolore magna aliqua. ';
const L22 = 'Lorem ipsum dolor sit elit, sed do eiusmod';
const L23 = 'Lorem dolor sit amet, consectetur elit';
const L24 = 'Lorem ipsum dolor sit, consectetur elit';
const L25 = 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod';
const L26 = 'www.yourgreatsite.com';
const L27 = 'Riverside 16, Malang';
const L28 = 'Start Presentation';
const L29 = 'Artificial Intelligence ';
const L30 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis.';
const L31 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore';
const L32 = 'Lorem ipsum dolor amet, consectetur adipiscing';
const L33 = 'Lorem ipsum dolor sit amet, consectetur sed do eiusmod labore';
const L34 = 'Through Intelligent AI Solutions';
const L35 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis';
const L36 = 'Lorem ipsum dolor sit amet, elit, ';
const L37 = 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut';
const L38 = 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim.';
const L39 = 'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor incididunt ut et dolore magna aliqua.';
const L40 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt';
const L41 = 'Lorem ipsum dolor sit amet, consectetur elit, sed do';
const L42 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi';
const L43 = 'Lorem ipsum dolor ';
const L44 = 'amet, consectetur adipiscing elit, ';
const L45 = 'www.yourwebsite.com';
function slide1(s) {          // Cover Slide A01
  photo(s, 7.78, 1.44, 3.74, 4.79);
  photo(s, 7.4, 3.37, 0.76, 0.76, 0.38);
  photo(s, 11.14, 2, 0.76, 0.76, 0.38);
  txt(s, 1.54, 1.47, 5.61, 2.12, [['How ', INK], ['AI Technology ', TEAL], ['is Transforming the World', INK]], {sz: 40});
  txt(s, 1.54, 3.59, 4.57, 0.34, L14, {c: GREY});
  body(s, 1.54, 4.83, 4.57, 0.68, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et.');
  pill(s, 1.61, 5.64, 1.73, 0.41, 'How It Works', {sz: 11, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.2});
  pill(s, 3.43, 5.64, 1.84, 0.41, L28, {sz: 11, al: 'c', va: 'm', line: TEAL, shadow: 1, r: 0.2});
  card(s, 6.71, 4.29, 1.73, 1.78, INDIGO, 0.16);
  txt(s, 6.89, 5.08, 1.45, 0.81, L15, {c: W});
  icon(s, 7.03, 4.51, 0.24, 0.34, W);
}
function slide2(s) {          // Cover Slide A02
  photo(s, 1.58, 1.38, 3.77, 4.78);
  card(s, 1.58, 4.38, 3.77, 1.78, INDIGO, 0.17);
  txt(s, 5.77, 1.47, 6.44, 1.92, [['AI for Everyone: Making ', INK], [L29, TEAL], ['Accessible', INK]], {sz: 36});
  body(s, 7.57, 5.04, 4.44, 0.98, L3);
  card(s, 5.6, 4.38, 1.73, 1.78, INDIGO, 0.16);
  txt(s, 5.77, 5.17, 1.45, 0.81, L15, {c: W});
  icon(s, 5.92, 4.6, 0.24, 0.34, W);
  oval(s, 1.76, 1.47, 0.53, 0.53, W, {shadow: 1});
  icon(s, 1.89, 1.62, 0.28, 0.22, TEAL);
  txt(s, 1.64, 2.08, 1.45, 0.57, 'Intelligence Evolution');
}
function slide3(s) {          // Cover Slide A03
  photo(s, 3.92, 3.64, 2.17, 2.41);
  photo(s, 8.55, 1.33, 3.19, 2.3);
  txt(s, 1.53, 1.38, 7.5, 1.45, [['The Role of AI in ', INK], ['Digital Transformation', TEAL]], {sz: 40});
  card(s, 7.05, 5.33, 1.43, 0.42, INDIGO, 0.21);
  txt(s, 7.05, 5.4, 1.43, 0.29, 'Join Us Now', {sz: 11, c: W, al: 'c'});
  body(s, 7, 4.23, 4.37, 0.98, L3);
  card(s, 9.03, 1.33, 2.72, 2.3, INDIGO, 0.18);
  card(s, 1.65, 3.64, 2.17, 2.41, INDIGO, 0.17);
  txt(s, 1.77, 5.02, 2.04, 0.81, 'The Power of AI in Digital Marketing & Advertising', {c: W});
  oval(s, 3.06, 3.87, 0.53, 0.53, W);
  icon(s, 3.19, 4.02, 0.28, 0.22, INDIGO);
}
function slide4(s) {          // Cover Slide A04
  photo(s, 7.73, 1.28, 4.02, 4.94);
  photo(s, 3.7, 2.81, 0.53, 0.53, 0.27);
  photo(s, 4.09, 2.81, 0.53, 0.53, 0.27);
  txt(s, 1.58, 1.54, 7.24, 1.92, [['Chatbots and ', INK], ['AI Assistants: ', TEAL], ['The Future ', INK], ['of Customer Service', INK]], {sz: 36});
  body(s, 1.58, 4.98, 4.87, 0.98, L16);
  card(s, 8.47, 3.81, 3.29, 2.41, INDIGO, 0.19);
  txt(s, 4.69, 2.78, 0.71, 0.4, [['16', INK], ['K', TEAL]]);
  body(s, 4.69, 3, 1.16, 0.38, 'Innovator');
}
function slide5(s) {          // Cover Slide A05
  photo(s, 0, 0.6, 6.67, 6.9, 0);
  txt(s, 7, 1.35, 5.23, 2.52, [['How AI is ', INK], ['Revolutionizing Disease ', TEAL], ['Diagnosis & Treatment', INK]], {sz: 36});
  body(s, 7, 4.64, 5, 0.98, L16);
  pill(s, 7.09, 5.69, 1.73, 0.41, 'Start Now', {sz: 11, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.2});
  card(s, 1.7, 2.47, 1.48, 1.35, INDIGO, 0.16);
  txt(s, 1.85, 2.61, 1.22, 0.5, '15K', {sz: 24, c: W});
  body(s, 1.85, 3, 1.48, 0.68, L17, W);
}
function slide6(s) {          // Cover Slide A06
  photo(s, 7.08, 3.05, 4.66, 2.96, 0.23);
  body(s, 1.53, 4.32, 4.72, 0.98, L30);
  txt(s, 1.51, 1.46, 8.49, 1.45, [['How AI Tech is', INK], [' Enhancing the Learning', TEAL], [' Experience', INK]], {sz: 40});
  pill(s, 1.61, 5.39, 1.73, 0.41, 'Start Now', {sz: 11, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.2});
  card(s, 10.25, 1.58, 1.48, 1.35, INDIGO, 0.16);
  txt(s, 10.41, 1.72, 1.22, 0.5, '35K', {sz: 24, c: W});
  body(s, 10.41, 2.11, 1.48, 0.68, L17, W);
  txt(s, 7.39, 2.23, 2.27, 0.57, L14);
}
function slide7(s) {          // Cover Slide A07
  photo(s, 6.84, 3.05, 4.88, 3.03, 0.18);
  photo(s, 1.7, 2.78, 0.52, 0.52, 0.26);
  photo(s, 2.06, 2.78, 0.52, 0.52, 0.26);
  txt(s, 1.58, 1.51, 7.03, 1.31, 'The Role of AI in Upskilling & Workforce Training', {sz: 36});
  body(s, 1.6, 5.04, 4.88, 0.98, L16);
  txt(s, 2.63, 2.74, 0.71, 0.4, [['28', INK], ['K', TEAL]]);
  body(s, 2.63, 2.96, 1.16, 0.38, 'Innovator');
  card(s, 9.7, 1.56, 1.73, 1.78, INDIGO, 0.16);
  txt(s, 9.88, 2.35, 1.45, 0.81, L15, {c: W});
  icon(s, 10.02, 1.78, 0.24, 0.34, W);
}
function slide8(s) {          // Cover Slide A08
  photo(s, 1.7, 3.86, 4.96, 2.52);
  txt(s, 1.76, 1.32, 7.17, 1.92, [['AI-Powered Cybersecurity: ', INK], ['The Next Frontier ', TEAL], ['in Digital Defense', INK]], {sz: 36});
  txt(s, 4.16, 2.59, 2.69, 0.57, 'AI Technology Presentation');
  pill(s, 9.24, 1.42, 2.44, 0.57, 'Investment Wins', {sz: 11, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.29});
  pill(s, 9.24, 2.05, 2.44, 0.57, 'Techniques', {sz: 11, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.29});
  pill(s, 9.24, 2.68, 2.44, 0.57, 'Stock Analysis', {sz: 11, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.29});
  body(s, 7.11, 4.43, 4.72, 0.98, L30);
  pill(s, 7.19, 5.5, 1.73, 0.41, 'Start Now', {sz: 11, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.2});
}
function slide9(s) {          // Cover Slide A09
  photo(s, 7.14, 0.71, 4.92, 5.68);
  card(s, 8.35, 2.1, 3.71, 4.28, INDIGO, 0.21);
  txt(s, 1.58, 1.33, 6.49, 1.92, [['AI-Generated Writing', TEAL], [': Can Machines Be Creative?', INK]], {sz: 36});
  txt(s, 1.58, 5.81, 2.27, 0.57, L14);
  oval(s, 3.59, 5.88, 0.51, 0.51, INDIGO);
  body(s, 1.58, 3.36, 4.43, 0.68, L31);
  card(s, 6.19, 4.42, 1.48, 1.35, INDIGO, 0.16);
  txt(s, 6.35, 4.56, 1.22, 0.5, '35K', {sz: 24, c: W});
  body(s, 6.35, 4.95, 1.48, 0.68, L17, W);
}
function slide10(s) {          // Cover Slide A10
  photo(s, 1.6, 1.46, 4.02, 3.56, 0.24);
  txt(s, 6.15, 1.46, 5.93, 1.92, [['Superintelligence: ', INK, {br: 1}], ['Will ', INK], ['AI Surpass Human ', TEAL], ['Intelligence?', INK]], {sz: 36});
  body(s, 7.31, 5.02, 4.43, 0.98, L3);
  txt(s, 6.15, 4.53, 5.52, 0.34, 'The Artificial Intelligent Technology Presentation');
  oval(s, 6.24, 5.14, 0.8, 0.8, INDIGO);
  icon(s, 6.5, 5.35, 0.29, 0.42, W);
  pill(s, 9.57, 2.8, 1.73, 0.41, L28, {sz: 11, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.2});
  card(s, 1.6, 4.71, 4.02, 1.33, INDIGO, 0.21, {top: 1});
  txt(s, 1.85, 5.04, 1.18, 0.4, '490K', {c: W});
  txt(s, 3.17, 5.04, 1.18, 0.4, [['4.8', W], ['/5', W, {sz: 14}]]);
  txt(s, 4.49, 5.04, 1.18, 0.4, '98%', {c: W});
  body(s, 1.85, 5.37, 1.18, 0.38, 'Clients', W);
  body(s, 4.49, 5.37, 1.18, 0.38, 'Satisfied', W);
  body(s, 3.17, 5.37, 1.18, 0.38, 'Best Rated', W);
}
function slide11(s) {          // About Us A01
  photo(s, 1.72, 1.36, 2.51, 2.73, 0.24);
  txt(s, 4.66, 1.42, 7.42, 1.31, [['Where ', INK], [L29, TEAL], ['Meets Business Excellence', INK]], {sz: 36});
  txt(s, 1.72, 4.59, 2.93, 0.34, 'AI-Driven Innovation');
  card(s, 7.56, 4.1, 1.99, 2.05, INDIGO, 0.19);
  txt(s, 8.14, 4.29, 1.38, 0.5, L10, {sz: 12, c: W});
  card(s, 9.66, 4.1, 1.99, 2.05, W, 0.17, {shadow: 1});
  txt(s, 10.32, 4.29, 1.3, 0.5, 'Customer Aspirations.', {sz: 12});
  icon(s, 7.76, 4.37, 0.35, 0.35, W);
  icon(s, 9.87, 4.37, 0.35, 0.35, TEAL);
  body(s, 1.72, 4.95, 4.43, 0.98, L3);
  body(s, 7.73, 4.95, 1.83, 0.98, L32, W);
  body(s, 9.84, 4.95, 1.83, 0.98, L32);
  body(s, 4.66, 2.74, 6.14, 0.68, L3);
}
function slide12(s) {          // About Us A02
  photo(s, 6.55, 1.3, 5.15, 2.49);
  txt(s, 1.92, 1.68, 4.21, 1.45, [['About Our ', INK], ['Company', TEAL]], {sz: 40});
  body(s, 6.78, 5.31, 4.69, 0.68, L18);
  body(s, 6.78, 4.25, 4.69, 0.98, L3);
  card(s, 1.67, 3.89, 2.33, 2.34, INDIGO, 0.15);
  txt(s, 1.92, 4.18, 1.94, 0.57, '24K+', {sz: 28, c: W});
  body(s, 1.92, 5, 1.95, 0.98, L33, W);
  card(s, 4.1, 3.89, 2.33, 2.34, W, 0.15, {shadow: 1});
  txt(s, 4.35, 4.18, 1.94, 0.57, [['62', INK], ['K', TEAL]], {sz: 28});
  body(s, 4.35, 5, 1.95, 0.98, L33);
}
function slide13(s) {          // About Us A03
  photo(s, 6.82, 1.43, 4.76, 3.04, 0.21);
  card(s, 6.82, 4.58, 4.76, 1.54, INDIGO, 0.18);
  txt(s, 1.75, 1.58, 4.69, 1.45, [['Introduction', TEAL], [' to Our Company', INK]], {sz: 40});
  body(s, 7.09, 4.81, 4.44, 0.68, L18, W);
  txt(s, 7.43, 5.6, 1.64, 0.3, 'The Winner', {sz: 12, c: W});
  oval(s, 7.18, 5.65, 0.22, 0.22, W);
  txt(s, 9.53, 5.6, 1.89, 0.3, 'Good Reputation', {sz: 12, c: W});
  oval(s, 9.28, 5.65, 0.22, 0.22, W);
  body(s, 1.75, 5.09, 4.69, 0.98, L19);
  txt(s, 1.75, 4.26, 1.18, 0.4, [['490', INK], ['K', TEAL]]);
  txt(s, 3.18, 4.26, 1.18, 0.4, [['4.8', INK], ['/5', TEAL, {sz: 14}]]);
  txt(s, 4.62, 4.26, 1.18, 0.4, [['98', INK], ['%', TEAL]]);
  body(s, 1.75, 4.58, 1.18, 0.36, 'Clients');
  body(s, 4.62, 4.58, 1.18, 0.36, 'Satisfied');
  body(s, 3.18, 4.59, 1.18, 0.36, 'Best Rated');
}
function slide14(s) {          // About Us A04
  photo(s, 7.89, 1.4, 3.72, 4.87, 0.21);
  txt(s, 1.68, 1.4, 6.84, 2.2, [['Transforming Ideas into ', INK], ['AI-Powered Realities', TEAL]], {sz: 40, lp: 50});
  body(s, 1.68, 4.73, 4.59, 0.98, L19);
  card(s, 6.67, 4.31, 2.01, 1.57, INDIGO, 0.19);
  txt(s, 1.68, 4.32, 3.85, 0.34, 'Smart AI Technology');
  txt(s, 6.79, 5.16, 1.86, 0.57, L34, {sz: 12, c: W, lp: 17.4});
  icon(s, 6.88, 4.54, 0.34, 0.34, W);
  oval(s, -0.44, 6.69, 1.62, 1.62, null, {line: HAIR});
  oval(s, -0.68, 6.45, 2.1, 2.1, null, {line: HAIR});
  oval(s, -0.89, 6.23, 2.53, 2.53, null, {line: HAIR});
  pill(s, 1.73, 5.86, 1.73, 0.41, 'Start Now', {sz: 11, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.2});
}
function slide15(s) {          // About Us A05
  photo(s, 3.86, 3.07, 5.62, 2.25);
  txt(s, 1.86, 1.28, 9.62, 1.45, [['Your ', INK], ['Trusted Partner ', TEAL], ['in AI and Automation', INK]], {sz: 40, al: 'c'});
  card(s, 1.45, 3.07, 2.3, 2.25, INDIGO, 0.16);
  card(s, 9.58, 3.07, 2.3, 2.25, W, 0.16, {shadow: 1});
  txt(s, 1.86, 5.66, 8.66, 0.68, L19, {sz: 12, f: BODY, c: GREY, al: 'c', ls: 1.5});
  txt(s, 2.14, 3.36, 1.38, 0.5, L10, {sz: 12, c: W});
  icon(s, 1.67, 3.41, 0.37, 0.37, W);
  body(s, 1.57, 4.12, 2.13, 0.98, L20, W);
  txt(s, 10.27, 3.36, 1.38, 0.5, L10, {sz: 12});
  body(s, 9.7, 4.12, 2.13, 0.98, L20);
  icon(s, 9.88, 3.47, 0.3, 0.26, TEAL);
}
function slide16(s) {          // About Us A06
  photo(s, 3.61, 1.14, 3.06, 2.62);
  photo(s, 1.54, 3.84, 2.72, 2.55);
  txt(s, 7.22, 1.5, 5.07, 1.92, [['Turning Data into Decisions with ', INK], ['AI Intelligence', TEAL]], {sz: 36});
  body(s, 7.22, 4.91, 4.74, 0.97, L35);
  card(s, 1.54, 1.67, 2.84, 0.83, INDIGO, 0.14);
  icon(s, 1.7, 1.87, 0.42, 0.42, W);
  txt(s, 2.12, 1.8, 2.28, 0.57, 'Precision, Passion, Performance', {c: W});
  card(s, 4.37, 3.84, 2.3, 2.25, INDIGO, 0.16);
  txt(s, 5.05, 4.13, 1.38, 0.5, L10, {sz: 12, c: W});
  icon(s, 4.58, 4.18, 0.37, 0.37, W);
  body(s, 4.48, 4.89, 2.13, 0.98, L20, W);
}
function slide17(s) {          // About Us A07
  photo(s, 6.99, 3.51, 2.3, 2.72);
  photo(s, 9.41, 3.51, 2.3, 2.72);
  txt(s, 1.56, 1.29, 6.39, 1.92, [['Optimizing ', INK], ['Business Performance ', TEAL], ['with Automation', INK]], {sz: 36});
  body(s, 1.56, 4.29, 4.74, 0.68, L6);
  txt(s, 1.88, 5.14, 3.62, 0.3, 'Intelligent Automation Implementation', {sz: 12});
  oval(s, 1.66, 5.21, 0.17, 0.17, INDIGO);
  txt(s, 1.88, 5.5, 3.62, 0.3, 'AI-powered Chatbot Integration', {sz: 12});
  oval(s, 1.66, 5.56, 0.17, 0.17, INDIGO);
  txt(s, 1.88, 5.81, 3.62, 0.3, 'Automated Ticketing Systems', {sz: 12});
  oval(s, 1.66, 5.88, 0.17, 0.17, INDIGO);
  body(s, 7.69, 1.95, 4.1, 0.98, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim.');
  txt(s, 7.69, 1.5, 1.38, 0.44, '60%', {sz: 20, c: TEAL});
}
function slide18(s) {          // About Us A08
  photo(s, 1.77, 2.22, 4.74, 3.82, 0.19);
  txt(s, 7.1, 1.64, 4.69, 1.45, [['About ', INK], ['Our Company', TEAL]], {sz: 40});
  body(s, 7.09, 3.86, 4.69, 0.98, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud');
  body(s, 8.02, 4.93, 3.69, 0.98, 'Lorem ipsum dolor sit amet, consectetur  elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad');
  oval(s, 7.12, 5.1, 0.71, 0.71, INDIGO);
  card(s, 4.24, 1.38, 2.01, 1.57, INDIGO, 0.19);
  txt(s, 4.35, 2.22, 1.86, 0.57, L34, {sz: 12, c: W, lp: 17.4});
  icon(s, 4.44, 1.6, 0.34, 0.34, W);
  card(s, 2.14, 1.38, 2.01, 1.57, W, 0.19, {shadow: 1});
  txt(s, 2.25, 2.22, 2.01, 0.57, 'Through  Industries Worldwide', {sz: 12, lp: 17.4});
  icon(s, 2.34, 1.6, 0.34, 0.34, TEAL);
}
function slide19(s) {          // About Us A09
  photo(s, 6.95, 1.31, 4.78, 3.08, 0.25);
  txt(s, 1.72, 1.69, 4.64, 1.92, [['We Are ', INK], ['A Global Leader ', TEAL], ['in AI-Powered', INK]], {sz: 36});
  body(s, 7.3, 5.07, 4.42, 0.97, L3);
  txt(s, 7.3, 4.68, 3.49, 0.34, 'Bring Intelligent AI Solution');
  card(s, 1.55, 4.54, 2.57, 1.7, INDIGO, 0.19, {shadow: 1});
  card(s, 4.24, 4.54, 2.57, 1.7, W, 0.19, {shadow: 1});
  body(s, 4.39, 5.28, 2.31, 0.68, L7);
  txt(s, 4.39, 4.8, 2.24, 0.5, [['182', INK], ['+', TEAL]], {sz: 24});
  body(s, 1.72, 5.28, 2.35, 0.68, L7, W);
  txt(s, 1.72, 4.8, 2.1, 0.5, '35M', {sz: 24, c: W});
}
function slide20(s) {          // About Us A10
  photo(s, 1.67, 3.7, 4.85, 2.63, 0.19);
  card(s, 10.18, 1.66, 1.34, 1.11, INDIGO, 0.08);
  txt(s, 1.82, 1.37, 7.88, 1.58, [['AI Partner ', TEAL], ['for Businesses Around the Globe', INK]], {sz: 44});
  body(s, 6.91, 3.86, 4.75, 0.68, L6);
  txt(s, 7.49, 4.83, 3.19, 0.34, 'The Foundation of Expertise', {c: INK2});
  txt(s, 7.49, 5.6, 3.19, 0.34, 'Driven by Data', {c: INK2});
  txt(s, 10.19, 1.8, 1.31, 0.81, 'The Biggest In Town', {c: W, al: 'c'});
  oval(s, 6.97, 5.68, 0.47, 0.47, INDIGO);
  icon(s, 7.12, 5.82, 0.17, 0.19, W);
  oval(s, 6.97, 4.91, 0.47, 0.47, INDIGO);
  icon(s, 7.1, 5.03, 0.19, 0.24, W);
  body(s, 7.51, 5.08, 4.15, 0.38, L7);
  body(s, 7.51, 5.84, 4.15, 0.38, L7);
}
function slide21(s) {          // Agenda I05
  photo(s, 1.38, 2.16, 10.58, 3.18, 0.22);
  txt(s, 2.72, 1.03, 7.89, 0.71, [['Today’s ', INK], ['Discussion Points', TEAL]], {sz: 36, al: 'c'});
  card(s, 1.56, 4.23, 3.19, 2.29, INDIGO, 0.16);
  txt(s, 2.56, 4.54, 2.25, 0.37, 'First Agenda ', {sz: 16, c: W});
  oval(s, 1.92, 4.47, 0.51, 0.51, W);
  body(s, 1.85, 5.19, 2.64, 0.98, L21, W);
  txt(s, 1.85, 4.54, 0.66, 0.37, '01.', {sz: 16, al: 'c'});
  card(s, 8.56, 4.23, 3.19, 2.29, INDIGO, 0.17);
  txt(s, 9.56, 4.54, 2.25, 0.37, 'Third Agenda ', {sz: 16, c: W});
  oval(s, 8.92, 4.47, 0.51, 0.51, W);
  body(s, 8.85, 5.19, 2.64, 0.98, L21, W);
  txt(s, 8.85, 4.55, 0.66, 0.37, '03.', {sz: 16, al: 'c'});
  card(s, 5.07, 4.23, 3.19, 2.29, W, 0.17, {shadow: 1});
  oval(s, 5.45, 4.47, 0.51, 0.51, TEAL);
  body(s, 5.38, 5.19, 2.64, 0.98, L21);
  txt(s, 5.38, 4.54, 0.66, 0.37, '02.', {sz: 16, c: W, al: 'c'});
  txt(s, 6.16, 4.54, 2.25, 0.37, 'Second Agenda', {sz: 16});
}
function slide22(s) {          // Agenda I06
  photo(s, 1.89, 1.22, 3.78, 4.77, 0.2);
  card(s, 4.15, 3.59, 2.64, 2.65, INDIGO, 0.24);
  txt(s, 4.44, 3.87, 1.79, 0.77, [['24,', W], [' June', W, {sz: 16}]], {sz: 40});
  txt(s, 4.91, 4.88, 1.87, 0.34, 'Main Office Hall', {c: W});
  txt(s, 4.91, 5.52, 1.87, 0.34, '08:00 - End', {c: W});
  icon(s, 4.46, 5.49, 0.39, 0.39, W);
  icon(s, 4.44, 4.88, 0.42, 0.42, W);
  txt(s, 7.29, 1.47, 4.86, 1.45, [['Our ', INK], ['Agenda', TEAL], [' to Success', INK]], {sz: 40});
  icon(s, 7.41, 5.28, 0.26, 0.26, TEAL);
  body(s, 7.29, 4.02, 4.58, 0.98, L2);
  txt(s, 7.78, 5.25, 1.99, 0.34, 'First Theme');
  icon(s, 7.41, 5.73, 0.26, 0.26, TEAL);
  txt(s, 7.78, 5.69, 1.99, 0.34, 'Third Theme');
  icon(s, 9.78, 5.28, 0.26, 0.26, TEAL);
  txt(s, 10.15, 5.25, 2.37, 0.34, 'Second Theme');
}
function slide23(s) {          // Agenda I07
  photo(s, 6.03, 1.18, 5.41, 2.52);
  txt(s, 1.59, 1.42, 4.18, 2.12, [['Agenda ', INK], ['Discussion ', TEAL], ['Calendar', INK]], {sz: 40});
  body(s, 1.59, 5.1, 3.43, 0.97, L6);
  txt(s, 1.59, 4.49, 3.19, 0.34, 'Guide to Our Strategy');
  card(s, 6.03, 3.81, 2.64, 2.49, INDIGO, 0.17);
  txt(s, 6.74, 5.2, 1.93, 0.34, 'Main Office Hall', {c: W});
  txt(s, 6.74, 5.71, 1.93, 0.34, '08:00 - End', {c: W});
  icon(s, 6.22, 5.67, 0.38, 0.38, W);
  icon(s, 6.21, 5.2, 0.42, 0.42, W);
  oval(s, 6.22, 4.07, 0.79, 0.79, W);
  txt(s, 6.12, 4.18, 0.99, 0.57, '21', {sz: 28, al: 'c'});
  txt(s, 7.19, 4.33, 1.24, 0.34, 'June 2026', {c: W});
  card(s, 8.81, 3.81, 2.64, 2.49, W, 0.17, {shadow: 1});
  txt(s, 9.51, 5.2, 1.93, 0.34, 'Main Office Hall');
  txt(s, 9.51, 5.71, 1.93, 0.34, '08:00 - End');
  icon(s, 9, 5.67, 0.38, 0.38, TEAL);
  icon(s, 8.98, 5.2, 0.42, 0.42, TEAL);
  oval(s, 8.99, 4.07, 0.79, 0.79, TEAL);
  txt(s, 8.89, 4.18, 0.99, 0.57, '26', {sz: 28, c: W, al: 'c'});
  txt(s, 9.96, 4.33, 1.24, 0.34, 'June 2026');
}
function slide24(s) {          // Agenda I08
  photo(s, 9.29, 2.72, 2.46, 3.38);
  photo(s, 6.66, 1.31, 2.46, 4.78, 0.18);
  body(s, 1.5, 4.2, 4.3, 0.97, L2);
  icon(s, 1.61, 5.48, 0.19, 0.14, TEAL);
  icon(s, 1.61, 5.9, 0.19, 0.14, TEAL);
  txt(s, 1.95, 5.38, 2.22, 0.3, 'Maintenance', {sz: 12});
  txt(s, 1.95, 5.79, 2.22, 0.3, 'Safety Standards', {sz: 12});
  txt(s, 1.5, 1.31, 4.99, 2.12, [['Today’s ', INK], ['Discussion', TEAL], [' Outline', INK]], {sz: 40});
  icon(s, 4, 5.48, 0.19, 0.14, TEAL);
  icon(s, 4, 5.9, 0.19, 0.14, TEAL);
  txt(s, 4.34, 5.38, 2.22, 0.3, 'Industry Landscape', {sz: 12});
  txt(s, 4.34, 5.79, 2.22, 0.3, 'Success Stories', {sz: 12});
  card(s, 9.62, 1.31, 1.81, 1.88, INDIGO, 0.21);
  txt(s, 9.62, 1.49, 1.81, 0.34, 'June, 2024', {c: W, al: 'c'});
  txt(s, 9.88, 2.63, 1.29, 0.34, 'Saturday', {f: BODY, c: W, al: 'c'});
  txt(s, 9.92, 1.81, 1.21, 0.91, '24', {sz: 48, c: W, al: 'c'});
}
function slide25(s) {          // Agenda I10
  photo(s, 8.04, 1.16, 3.73, 5.22, 0.23);
  txt(s, 1.67, 1.26, 5.85, 1.45, [['Agenda ', INK], ['Discussion Calendar', TEAL]], {sz: 40});
  card(s, 1.56, 4.33, 2.37, 1.88, INDIGO, 0.21);
  txt(s, 1.67, 4.95, 1.77, 0.34, 'Saturday', {f: BODY, c: W});
  txt(s, 1.67, 4.39, 1.29, 0.64, '22', {sz: 32, c: W});
  body(s, 1.67, 5.34, 2.24, 0.67, L22, W);
  card(s, 4.07, 4.33, 2.37, 1.88, INDIGO, 0.21);
  txt(s, 4.19, 4.95, 1.77, 0.34, 'Monday', {f: BODY, c: W});
  txt(s, 4.19, 4.39, 1.29, 0.64, '24', {sz: 32, c: W});
  body(s, 4.19, 5.34, 2.24, 0.67, L22, W);
  card(s, 6.59, 4.33, 2.37, 1.88, W, 0.21, {shadow: 1});
  txt(s, 6.7, 4.95, 1.77, 0.34, 'Tuesday', {f: BODY});
  txt(s, 6.7, 4.39, 1.29, 0.64, '25', {sz: 32, c: TEAL});
  body(s, 6.7, 5.34, 2.24, 0.67, L22);
  body(s, 1.67, 2.9, 5.53, 0.68, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ');
}
function slide26(s) {          // Multipurpose B01
  photo(s, 1.62, 3.93, 5.66, 2.23, 0.19);
  photo(s, 7.53, 1.2, 4.19, 2.49, 0.22);
  card(s, 9.75, 3.93, 1.97, 2.23, INDIGO, 0.21);
  txt(s, 1.75, 1.38, 5.82, 1.18, [['The Evolution of ', INK], ['Global Technology', TEAL]], {sz: 32});
  body(s, 1.75, 2.66, 5.32, 0.68, 'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim');
  card(s, 7.53, 3.91, 1.97, 2.23, W, 0.21, {shadow: 1});
  txt(s, 7.7, 5.22, 1.62, 0.68, L36, {sz: 12, f: BODY, c: GREY, al: 'c', ls: 1.5});
  txt(s, 7.57, 4.79, 1.87, 0.34, 'Hi Tech', {al: 'c'});
  txt(s, 9.93, 5.22, 1.62, 0.68, L36, {sz: 12, f: BODY, c: W, al: 'c', ls: 1.5});
  txt(s, 9.8, 4.79, 1.87, 0.34, 'Digital', {c: W, al: 'c'});
  oval(s, 8.25, 4.1, 0.51, 0.51, TEAL);
  oval(s, 10.48, 4.1, 0.51, 0.51, W);
  icon(s, 10.58, 4.21, 0.3, 0.3);
  icon(s, 8.36, 4.2, 0.31, 0.31, W);
}
function slide27(s) {          // Multipurpose B02
  photo(s, 1.66, 3.65, 5.27, 2.5, 0.26);
  txt(s, 1.73, 1.35, 5.27, 1.92, [['Shaping Economic Policies of ', INK], ['Global Technology', TEAL]], {sz: 36});
  body(s, 7.46, 1.32, 4.26, 0.98, L2);
  body(s, 7.46, 2.44, 4.26, 0.68, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ');
  card(s, 7.35, 5.56, 4.38, 0.59, W, 0.06, {shadow: 1});
  card(s, 7.35, 4.22, 4.38, 0.59, INDIGO, 0.06);
  txt(s, 7.41, 4.33, 0.53, 0.34, '01.', {c: W, b: 1});
  body(s, 7.89, 4.3, 4.09, 0.38, L23, W);
  card(s, 7.35, 4.89, 4.38, 0.59, W, 0.06, {shadow: 1});
  txt(s, 7.41, 5, 0.53, 0.34, '02.', {c: TEAL, b: 1});
  body(s, 7.89, 4.96, 4.09, 0.38, L11, INK);
  txt(s, 7.41, 5.67, 0.53, 0.34, '03.', {c: TEAL, b: 1});
  body(s, 7.89, 5.63, 4.09, 0.38, L24, INK);
}
function slide28(s) {          // Multipurpose B03
  photo(s, 7.34, 1.47, 4.32, 2.36);
  txt(s, 1.87, 4.58, 5.22, 1.31, [['The ', INK], ['Key to Thriving ', TEAL], ['in Technology', INK]], {sz: 36});
  card(s, 4.51, 1.47, 2.66, 2.36, INDIGO, 0.17);
  body(s, 4.7, 2.57, 2.46, 0.98, L1, W);
  oval(s, 4.74, 1.68, 0.59, 0.59, W);
  card(s, 1.68, 1.47, 2.66, 2.36, W, 0.17, {shadow: 1});
  body(s, 1.87, 2.57, 2.46, 0.98, L1);
  oval(s, 1.92, 1.68, 0.59, 0.59, TEAL);
  icon(s, 4.82, 1.75, 0.44, 0.44);
  icon(s, 1.99, 1.77, 0.44, 0.44, W);
  body(s, 7.53, 4.58, 4.26, 0.98, L2);
  icon(s, 7.63, 5.86, 0.19, 0.14, TEAL);
  txt(s, 7.83, 5.76, 2.22, 0.3, 'Maintenance', {sz: 12});
  icon(s, 9.59, 5.86, 0.19, 0.14, TEAL);
  txt(s, 9.79, 5.76, 2.22, 0.3, 'Landscape', {sz: 12});
}
function slide29(s) {          // Multipurpose B04
  photo(s, 1.57, 4.07, 2.78, 2.29, 0.22);
  photo(s, 7.78, 1.31, 3.98, 2.43);
  txt(s, 1.52, 1.58, 6.27, 0.71, [['AI Technology ', TEAL], ['Trends', INK]], {sz: 36});
  body(s, 1.52, 2.45, 5.6, 0.67, L12);
  card(s, 7.78, 4.07, 3.98, 2.29, INDIGO, 0.16, {shadow: 1});
  hline(s, 8.22, 5.15, 3.3, TEAL);
  hline(s, 8.22, 5.61, 3.3, TEAL);
  hline(s, 8.22, 6.07, 3.3, TEAL);
  card(s, 8.39, 5.08, 0.35, 0.99, W, 0.06, {top: 1});
  card(s, 9.05, 5.2, 0.35, 0.86, W, 0.06, {top: 1});
  card(s, 9.7, 5.36, 0.35, 0.71, W, 0.06, {top: 1});
  card(s, 10.36, 5.2, 0.35, 0.86, W, 0.06, {top: 1});
  card(s, 11.02, 5.08, 0.35, 0.99, W, 0.06, {top: 1});
  txt(s, 7.91, 5.49, 0.34, 0.24, '5', {sz: 8, c: W, al: 'c'});
  txt(s, 7.91, 5.95, 0.34, 0.24, '0', {sz: 8, c: W, al: 'c'});
  txt(s, 7.91, 5.04, 0.34, 0.24, '10', {sz: 8, c: W, al: 'c'});
  txt(s, 7.91, 4.28, 2.89, 0.34, 'Competitive Devaluation', {c: W});
  body(s, 4.88, 5.04, 2.38, 0.97, L1);
  txt(s, 4.88, 4.31, 2.38, 0.57, [['13', INK], ['K', TEAL]], {sz: 28});
}
function slide30(s) {          // Multipurpose B05
  photo(s, 1.68, 3.75, 2.19, 2.48);
  txt(s, 1.99, 1.58, 10.02, 1.45, [['The AI Revolution: ', INK], ['Navigating the ', TEAL], ['Age of Smart Technology', INK]], {sz: 40});
  oval(s, 6.83, 4.12, 0.53, 0.53, INDIGO);
  oval(s, 6.83, 5.2, 0.53, 0.53, INDIGO);
  body(s, 7.62, 3.98, 4.04, 0.98, L12);
  body(s, 7.62, 5.07, 4.04, 0.98, L12);
  card(s, 3.98, 3.75, 2.19, 2.48, INDIGO, 0.17);
  pill(s, 4.16, 5.62, 1.71, 0.42, 'Read More', {sz: 12, al: 'c', va: 'm', fill: W, r: 0.21});
  body(s, 4.19, 4.66, 1.97, 0.68, L11, W);
  icon(s, 6.94, 4.23, 0.31, 0.31, W);
  icon(s, 6.94, 5.33, 0.31, 0.31, W);
  oval(s, 4.27, 3.94, 0.53, 0.53, W);
  icon(s, 4.38, 4.05, 0.31, 0.31);
}
function slide31(s) {          // Multipurpose B06
  photo(s, 9.16, 4.07, 2.31, 2.47);
  card(s, 1.88, 4.07, 2.3, 2.47, W, 0.16, {shadow: 1});
  card(s, 1.88, 1.31, 2.31, 2.47, INDIGO, 0.16);
  txt(s, 4.79, 1.66, 7, 1.45, [['From Data to Decisions: ', INK], ['The Intelligence ', TEAL], ['of AI', INK]], {sz: 40});
  txt(s, 2.05, 2.5, 2.13, 1.04, 'Navigating Currency Markets: A Guide to Forex and Beyond', {c: W});
  oval(s, 3.45, 1.47, 0.56, 0.56, W);
  txt(s, 2.05, 5.36, 2.24, 0.81, 'Redefining Innovation and Intelligence', {c: INK});
  icon(s, 3.53, 1.55, 0.42, 0.42);
  oval(s, 3.44, 4.17, 0.56, 0.56, TEAL);
  body(s, 4.79, 5.19, 4.04, 0.98, L12);
  txt(s, 4.79, 4.42, 2.38, 0.57, [['13', INK], ['K', TEAL]], {sz: 28});
}
function slide32(s) {          // Multipurpose B07
  photo(s, 1.78, 1.43, 2.36, 2.47);
  photo(s, 6.48, 4.1, 2.55, 2.27);
  card(s, 9.19, 4.1, 2.31, 2.27, INDIGO, 0.16);
  txt(s, 4.62, 1.48, 7.19, 1.31, [['AI at Work', TEAL], [': Automating, Optimizing, and Innovating', INK]], {sz: 36});
  body(s, 1.65, 4.47, 4.38, 0.97, L2);
  txt(s, 9.38, 5.33, 2.01, 0.81, 'Identification Of Growth Areas And Market Gaps.', {c: W});
  oval(s, 9.48, 4.31, 0.56, 0.56, W);
  icon(s, 9.55, 4.38, 0.42, 0.42);
  pill(s, 1.7, 5.69, 1.62, 0.4, 'Read More', {sz: 12, f: BODY, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.2});
}
function slide33(s) {          // Multipurpose B08
  photo(s, 7.34, 2.57, 4.06, 3.88, 0.24);
  txt(s, 1.72, 1.25, 11.24, 1.31, [['AI-Powered Transformation: ', INK], ['Unlocking Business Potential', TEAL]], {sz: 36});
  body(s, 4.51, 5.11, 2.48, 0.98, L1);
  txt(s, 4.51, 4.7, 2.01, 0.34, 'Management ');
  card(s, 1.55, 3.96, 2.48, 2.48, INDIGO, 0.18);
  oval(s, 1.8, 4.21, 0.51, 0.51, W);
  txt(s, 1.72, 5.32, 2.26, 0.81, 'Ethical Innovation & Intelligent Great Solutions', {c: W});
  icon(s, 1.93, 4.34, 0.25, 0.25);
  body(s, 1.72, 2.72, 5.27, 0.68, L6);
}
function slide34(s) {          // Multipurpose B09
  photo(s, 3.47, 1.04, 8.14, 3.22, 0.25);
  body(s, 8.1, 5.37, 3.51, 0.98, L6);
  txt(s, 8.1, 4.86, 3.19, 0.34, 'Our General Purpose');
  txt(s, 1.88, 4.77, 6.19, 1.58, [['The Future of ', INK], ['AI Technology', TEAL]], {sz: 44});
  card(s, 1.69, 1.47, 2.29, 2.36, INDIGO, 0.17);
  txt(s, 1.88, 1.76, 1.92, 0.57, 'A Guide to Forex and Beyond', {c: W});
  txt(s, 1.88, 2.57, 2.1, 0.98, [['Lorem ipsum dolor sit amet, adipiscing elit, ', W, {br: 1}], ['do eiusmod', W]], {sz: 12, f: BODY, ls: 1.5});
}
function slide35(s) {          // Multipurpose B10
  photo(s, 6.77, 3.67, 4.95, 2.46, 0.19);
  body(s, 1.61, 5.02, 4.75, 0.98, L35);
  txt(s, 1.61, 4.45, 3.19, 0.34, 'The Future of Global Finance');
  txt(s, 1.61, 1.34, 5.66, 1.92, [['AI in Action: ', INK], ['Real-World Applications ', TEAL], ['Breakthroughs', INK]], {sz: 36});
  card(s, 7.35, 2.68, 4.38, 0.59, W, 0.06, {shadow: 1});
  card(s, 7.35, 1.34, 4.38, 0.59, INDIGO, 0.06);
  txt(s, 7.41, 1.45, 0.53, 0.34, '01.', {c: W, b: 1});
  body(s, 7.89, 1.42, 4.09, 0.38, L23, W);
  card(s, 7.35, 2.01, 4.38, 0.59, W, 0.06, {shadow: 1});
  txt(s, 7.41, 2.12, 0.53, 0.34, '02.', {c: TEAL, b: 1});
  body(s, 7.89, 2.08, 4.09, 0.38, L11, INK);
  txt(s, 7.41, 2.79, 0.53, 0.34, '03.', {c: TEAL, b: 1});
  body(s, 7.89, 2.75, 4.09, 0.38, L24, INK);
}
function slide36(s) {          // Our Service B01
  txt(s, 7.48, 1.57, 4.32, 2.12, [['The ', INK], ['Best Service', TEAL], [' You Get', INK]], {sz: 40});
  body(s, 7.48, 4.64, 3.98, 1.29, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam qu nostrud exercitation ullamco laboris nisi');
  txt(s, 7.48, 4.23, 1.69, 0.34, 'Best in Town');
  card(s, 1.6, 1.4, 2.47, 2.3, INDIGO, 0.13, {shadow: 1});
  txt(s, 1.79, 2.47, 2.28, 0.91, L1, {sz: 11, f: BODY, c: W, ls: 1.5});
  txt(s, 2.38, 1.66, 1.69, 0.57, 'Technology Hedging ', {c: W});
  oval(s, 1.83, 1.71, 0.48, 0.48, W);
  card(s, 4.2, 1.4, 2.47, 2.3, W, 0.13, {shadow: 1});
  oval(s, 4.41, 1.71, 0.48, 0.48, TEAL);
  txt(s, 4.37, 2.47, 2.28, 0.9, L1, {sz: 11, f: BODY, c: GREY, ls: 1.5});
  txt(s, 4.97, 1.65, 1.69, 0.57, 'Global Technology');
  card(s, 4.2, 3.81, 2.47, 2.3, INDIGO, 0.13);
  oval(s, 4.43, 4.12, 0.48, 0.48, W);
  txt(s, 4.39, 4.89, 2.25, 0.9, L1, {sz: 11, f: BODY, c: W, ls: 1.5});
  txt(s, 4.98, 4.06, 1.69, 0.57, 'Technology Analytics', {c: W});
  icon(s, 1.96, 1.86, 0.24, 0.18);
  icon(s, 4.53, 4.24, 0.28, 0.22);
  card(s, 1.59, 3.8, 2.47, 2.3, W, 0.13, {shadow: 1});
  oval(s, 1.81, 4.11, 0.48, 0.48, TEAL);
  txt(s, 1.76, 4.88, 2.28, 0.9, L1, {sz: 11, f: BODY, c: GREY, ls: 1.5});
  txt(s, 2.36, 4.05, 1.69, 0.57, 'Tech Development');
  icon(s, 1.91, 4.23, 0.29, 0.21, W);
}
function slide37(s) {          // Our Service B02
  photo(s, 4.98, 1.3, 3.04, 4.97, 0.19);
  card(s, 1.6, 2.88, 3.04, 1.81, INDIGO, 0.11);
  body(s, 1.76, 1.76, 2.79, 0.67, L5);
  body(s, 1.76, 3.66, 2.79, 0.67, L5, W);
  body(s, 1.76, 5.59, 2.79, 0.67, L5);
  txt(s, 1.76, 1.3, 2.17, 0.34, [['Our Service ', INK], ['#0', TEAL], ['1', INK]]);
  txt(s, 1.76, 3.17, 2.17, 0.34, 'Our Service #03', {c: W});
  txt(s, 1.76, 5.09, 2.17, 0.34, [['Our Service ', INK], ['#04', TEAL]]);
  txt(s, 8.32, 1.45, 3.77, 1.45, [['Service ', TEAL], ['We Are Offered', INK]], {sz: 40});
  txt(s, 8.32, 4.16, 3.77, 0.98, [['Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore ', GREY, {br: 1}], ['et dolore magna aliqua. Ut enim.', GREY]], {sz: 12, f: BODY, ls: 1.5});
  body(s, 8.32, 5.26, 3.77, 0.68, L37);
}
function slide38(s) {          // Our Service B05
  photo(s, 1.92, 1.37, 3.63, 4.81, 0.22);
  txt(s, 6.06, 1.42, 6.15, 1.45, [['About Our ', INK], ['Great Service', TEAL]], {sz: 40});
  body(s, 6.06, 5.11, 2.04, 0.98, L25);
  txt(s, 6.06, 4.54, 0.89, 0.57, '01.', {sz: 28, c: TEAL});
  body(s, 8.14, 5.11, 2.04, 0.98, L25);
  txt(s, 8.14, 4.54, 0.89, 0.57, '02.', {sz: 28, c: TEAL});
  body(s, 10.21, 5.11, 2.04, 0.98, L25);
  txt(s, 10.21, 4.54, 0.89, 0.57, '03.', {sz: 28, c: TEAL});
  body(s, 6.06, 2.95, 5.7, 0.68, L38);
  card(s, 1.55, 1.51, 1.23, 1.47, INDIGO, 0.12, {shadow: 1});
  txt(s, 1.64, 2.25, 1.05, 0.57, 'Best in Town', {c: W, al: 'c'});
  oval(s, 1.92, 1.71, 0.48, 0.48, W);
}
function slide39(s) {          // Our Service B06
  photo(s, -0.02, 0.93, 4.75, 5.65, 0);
  txt(s, 5.43, 1.45, 6.18, 1.45, [['About Our ', INK], ['Great Service', TEAL]], {sz: 40});
  card(s, 1.88, 4.42, 3.2, 1.64, W, 0.17, {shadow: 1});
  oval(s, 2.14, 4.6, 0.58, 0.58, TEAL);
  body(s, 2.07, 5.2, 2.83, 0.67, L5);
  txt(s, 2.86, 4.72, 2.03, 0.34, 'Service #01');
  card(s, 5.24, 4.42, 3.2, 1.64, INDIGO, 0.17);
  oval(s, 5.49, 4.6, 0.58, 0.58, W);
  body(s, 5.42, 5.2, 2.83, 0.68, L5, W);
  txt(s, 6.22, 4.72, 2.03, 0.34, 'Service #02', {c: W});
  card(s, 8.59, 4.42, 3.2, 1.64, INDIGO, 0.17, {shadow: 1});
  oval(s, 8.85, 4.6, 0.58, 0.58, W);
  body(s, 8.78, 5.2, 2.83, 0.67, L5, W);
  txt(s, 9.58, 4.72, 2.03, 0.34, 'Service #03', {c: W});
  icon(s, 8.97, 4.72, 0.35, 0.35);
  icon(s, 5.61, 4.72, 0.35, 0.35);
  icon(s, 2.26, 4.71, 0.33, 0.33, W);
  body(s, 5.43, 3.02, 5.7, 0.68, L38);
}
function slide40(s) {          // Our Service B07
  photo(s, 6.05, 1.37, 3.1, 4.81, 0.24);
  photo(s, 1.54, 4.45, 2.17, 1.73);
  txt(s, 1.8, 1.47, 3.75, 1.31, [['Our ', INK], ['Great Service', TEAL]], {sz: 36});
  body(s, 9.45, 1.91, 2.5, 0.97, L8);
  txt(s, 9.45, 1.49, 2.5, 0.34, [['Service ', INK], ['#01', TEAL, {b: 1}]]);
  body(s, 9.45, 3.53, 2.5, 0.98, L8);
  txt(s, 9.45, 3.11, 2.5, 0.34, [['Service ', INK], ['#02', TEAL, {b: 1}]]);
  body(s, 1.8, 2.92, 3.46, 0.68, L1);
  body(s, 9.45, 5.16, 2.5, 0.97, L8);
  txt(s, 9.45, 4.74, 2.5, 0.34, [['Service ', INK], ['#03', TEAL, {b: 1}]]);
  card(s, 3.8, 4.45, 2.17, 1.73, INDIGO, 0.13);
  body(s, 4.07, 5.06, 1.75, 0.97, L8, W);
  txt(s, 4.07, 4.64, 1.75, 0.34, 'Best in Town', {c: W});
}
function slide41(s) {          // Our Team B02
  photo(s, 6.67, 1.36, 1.91, 2.26);
  photo(s, 6.67, 3.98, 1.91, 2.26);
  txt(s, 1.56, 1.62, 4.88, 1.31, [['Team Role ', TEAL], ['& Responsibilities', INK]], {sz: 36});
  body(s, 1.56, 4.39, 4.29, 0.98, L2);
  body(s, 8.92, 1.88, 2.06, 0.37, 'Risk Management', INK);
  txt(s, 8.92, 1.64, 2.06, 0.34, 'Samsul Man', {c: TEAL});
  body(s, 8.92, 4.46, 2.06, 0.37, 'Cryptocurrency Trading', INK);
  txt(s, 8.92, 4.23, 2.06, 0.34, 'Michelle Ann', {c: TEAL});
  pill(s, 1.64, 5.51, 1.41, 0.4, 'Read More', {sz: 12, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.2});
  body(s, 8.92, 5.05, 2.99, 0.98, L39);
  body(s, 8.92, 2.42, 2.99, 0.98, L39);
}
function slide42(s) {          // Our Team B05
  photo(s, 9.18, 1.61, 2.47, 3.43);
  photo(s, 6.37, 2.49, 2.47, 3.43);
  txt(s, 1.49, 1.65, 3.95, 1.58, [['Meet ', INK], ['The Experts', TEAL]], {sz: 44});
  body(s, 1.53, 4.98, 4.1, 0.98, L2);
  txt(s, 1.53, 4.51, 2.53, 0.34, '10 Years of Experience');
  card(s, 6.37, 1.66, 2.47, 1.03, W, 0.14, {shadow: 1});
  txt(s, 6.37, 2.12, 2.47, 0.38, [['(', GREY], ['Profession A', GREY, {i: 1}], [')', GREY]], {sz: 12, f: BODY, al: 'c', ls: 1.5});
  txt(s, 6.37, 1.87, 2.47, 0.34, 'Fernandes Samuel', {c: TEAL, al: 'c'});
  card(s, 9.18, 4.84, 2.47, 1.03, INDIGO, 0.14, {shadow: 1});
  txt(s, 9.18, 5.3, 2.47, 0.38, [['(', W], ['Profession B', W, {i: 1}], [')', W]], {sz: 12, f: BODY, al: 'c', ls: 1.5});
  txt(s, 9.18, 5.05, 2.47, 0.34, 'Martin Fernandes', {c: W, al: 'c'});
}
function slide43(s) {          // Our Team B07
  photo(s, 4.05, 1.28, 3.21, 4.94);
  txt(s, 7.78, 1.74, 4.69, 1.31, [['Professional ', TEAL, {br: 1}], ['Who Help You', INK]], {sz: 36});
  body(s, 7.78, 3.83, 4.15, 0.98, L4);
  card(s, 2.31, 1.86, 1.42, 1.37, W, 0.09, {shadow: 1});
  txt(s, 2.41, 2.48, 1.44, 0.38, 'Michelle Jean', {sz: 12, ls: 1.5});
  oval(s, 2.5, 1.97, 0.43, 0.43, TEAL);
  txt(s, 2.43, 2.75, 1.42, 0.36, [['(Trader', GREY, {i: 1}], [')', GREY]], {sz: 11, f: BODY, ls: 1.5});
  card(s, 1.63, 3.37, 2.1, 2.27, INDIGO, 0.15, {shadow: 1});
  txt(s, 1.73, 4.6, 2.01, 0.81, 'Innovator of AI Technology for Education', {c: W});
  oval(s, 3.03, 3.52, 0.56, 0.56, W);
  icon(s, 3.11, 3.59, 0.42, 0.42);
  body(s, 7.78, 4.9, 4.15, 0.68, L40);
}
function slide44(s) {          // Our Team B09
  photo(s, 1.85, 3.66, 2.09, 1.84);
  photo(s, 4.36, 3.66, 2.09, 1.84);
  photo(s, 6.88, 1.42, 2.09, 1.84);
  photo(s, 9.4, 1.42, 2.09, 1.84);
  card(s, 6.74, 1.25, 2.37, 2.88, W, 0.19, {shadow: 1});
  card(s, 9.26, 1.25, 2.37, 2.88, INDIGO, 0.19, {shadow: 1});
  txt(s, 6.74, 3.44, 2.37, 0.34, 'George Fanny', {c: TEAL, al: 'c'});
  txt(s, 6.74, 3.7, 2.37, 0.29, 'Profession C', {sz: 11, f: BODY, c: GREY, al: 'c'});
  txt(s, 9.25, 3.44, 2.37, 0.34, 'Dean Fernando', {c: W, al: 'c'});
  txt(s, 9.25, 3.69, 2.37, 0.29, 'Profession D', {sz: 11, f: BODY, c: W, al: 'c'});
  txt(s, 1.85, 1.49, 3.84, 1.31, [['Introducing ', INK], ['Key Members', TEAL]], {sz: 36});
  card(s, 1.7, 3.49, 2.37, 2.88, W, 0.19, {shadow: 1});
  card(s, 4.22, 3.49, 2.37, 2.88, INDIGO, 0.19, {shadow: 1});
  body(s, 7.15, 5.27, 4.23, 0.98, L4);
  txt(s, 7.15, 4.77, 3.19, 0.34, 'Our Great Team');
  txt(s, 1.7, 5.68, 2.37, 0.34, 'Samantha Ann', {c: TEAL, al: 'c'});
  txt(s, 1.7, 5.93, 2.37, 0.29, 'Profession A', {sz: 11, f: BODY, c: GREY, al: 'c'});
  txt(s, 4.22, 5.68, 2.37, 0.34, 'Ronal Arvinando', {c: W, al: 'c'});
  txt(s, 4.22, 5.93, 2.37, 0.29, 'Profession B', {sz: 11, f: BODY, c: W, al: 'c'});
}
function slide45(s) {          // Our Team G01
  photo(s, 3.86, 1.12, 3.71, 5.28, 0.23);
  txt(s, 7.89, 1.75, 4.35, 1.45, [['Meet Our ', INK], ['Great Leader', TEAL]], {sz: 40});
  card(s, 1.62, 1.76, 2.63, 2.15, INDIGO, 0.24);
  body(s, 1.82, 2.98, 2.42, 0.68, L41, W);
  txt(s, 1.82, 2.61, 1.98, 0.34, 'Collaboration', {c: W});
  oval(s, 1.88, 1.97, 0.51, 0.51, W);
  txt(s, 1.88, 2.05, 0.51, 0.34, '01', {al: 'c'});
  card(s, 1.62, 4.04, 2.63, 2.15, W, 0.24, {shadow: 1});
  body(s, 1.86, 5.24, 2.42, 0.68, L41);
  txt(s, 1.86, 4.91, 1.98, 0.34, 'Achievement');
  oval(s, 1.9, 4.25, 0.51, 0.51, TEAL);
  txt(s, 1.9, 4.34, 0.51, 0.34, '02', {c: W, al: 'c'});
  body(s, 7.87, 4.77, 4.15, 0.98, L4);
  txt(s, 7.87, 4.33, 3.19, 0.34, 'Samantha Michelle');
}
function slide46(s) {          // Portfolio B01
  photo(s, 1.56, 3.29, 2.46, 3.23);
  photo(s, 6.74, 3.29, 2.46, 3.23);
  txt(s, 3.02, 1.08, 7.29, 0.84, [['Our Great ', INK], ['Showcase', TEAL]], {sz: 44, al: 'c'});
  txt(s, 2.29, 2.06, 8.77, 0.68, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris. ', {sz: 12, f: BODY, c: GREY, al: 'c', ls: 1.5});
  card(s, 4.14, 3.29, 2.48, 3.23, INDIGO, 0.2);
  card(s, 9.31, 3.29, 2.48, 3.23, W, 0.2, {shadow: 1});
  body(s, 4.33, 4.71, 2.26, 0.98, L13, W);
  txt(s, 4.33, 4.2, 1.98, 0.34, 'Project 01', {c: W});
  pill(s, 4.41, 5.9, 1.39, 0.34, 'Read More', {sz: 10, al: 'c', va: 'm', fill: W, r: 0.17});
  icon(s, 4.42, 3.61, 0.35, 0.38, W);
  body(s, 9.53, 4.71, 2.26, 0.98, L13);
  txt(s, 9.53, 4.2, 1.98, 0.34, 'Project 02');
  pill(s, 9.6, 5.9, 1.39, 0.34, 'Read More', {sz: 10, c: W, al: 'c', va: 'm', fill: TEAL, r: 0.17});
  icon(s, 9.6, 3.55, 0.36, 0.36, TEAL);
}
function slide47(s) {          // Portfolio B02
  photo(s, 1.59, 1.54, 4.7, 3.76, 0.23);
  txt(s, 7.05, 1.49, 5.26, 2.52, [['Highlight ', INK], ['The Exceptional ', TEAL], ['Showcase', INK]], {sz: 48});
  body(s, 7.05, 4.86, 4.11, 0.98, L4);
  card(s, 1.91, 4.98, 4.06, 0.98, INDIGO, 0.2);
  icon(s, 2.18, 5.27, 0.36, 0.36, W);
  body(s, 2.7, 5.17, 3.25, 0.68, 'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor incididunt ut', W);
}
function slide48(s) {          // Portfolio B04
  photo(s, 6.59, 3.83, 2.48, 2.55, 0.2);
  photo(s, 9.24, 3.83, 2.48, 2.55, 0.2);
  txt(s, 1.64, 1.37, 5.16, 1.72, [['Our Great ', INK], ['Showcase', TEAL]], {sz: 48});
  card(s, 6.59, 1.13, 2.48, 2.55, INDIGO, 0.2);
  card(s, 9.24, 1.13, 2.48, 2.55, W, 0.2, {shadow: 1});
  body(s, 6.79, 1.88, 2.26, 0.98, L13, W);
  txt(s, 6.79, 1.37, 1.98, 0.34, 'Project 01', {c: W});
  pill(s, 6.86, 3.06, 1.39, 0.34, 'Read More', {sz: 10, al: 'c', va: 'm', fill: W, r: 0.17});
  body(s, 9.46, 1.88, 2.26, 0.98, L13);
  txt(s, 9.46, 1.37, 1.98, 0.34, 'Project 02');
  pill(s, 9.53, 3.06, 1.39, 0.34, 'Read More', {sz: 10, c: W, al: 'c', va: 'm', fill: TEAL, r: 0.17});
  body(s, 1.64, 4.38, 4.15, 0.98, L4);
  body(s, 1.64, 5.45, 4.15, 0.68, L40);
}
function slide49(s) {          // Portfolio B06
  photo(s, 6.95, 1.28, 4.53, 3.98, 0.24);
  txt(s, 1.76, 1.53, 3.83, 1.45, [['Our ', INK], ['Portfolio ', TEAL], ['Showcase', INK]], {sz: 40});
  body(s, 1.76, 5.01, 4.09, 0.98, L4);
  txt(s, 1.76, 3.63, 3.19, 0.34, 'Portfolio 01');
  card(s, 1.85, 4.11, 0.73, 0.61, INDIGO, 0.14, {shadow: 1});
  card(s, 2.76, 4.11, 0.73, 0.61, W, 0.14, {shadow: 1});
  card(s, 3.67, 4.1, 0.73, 0.61, W, 0.14, {shadow: 1});
  card(s, 4.58, 4.11, 0.73, 0.61, W, 0.14, {shadow: 1});
  txt(s, 1.85, 4.21, 0.73, 0.4, '01', {c: W, al: 'c'});
  txt(s, 2.76, 4.22, 0.73, 0.4, '02', {c: TEAL, al: 'c'});
  txt(s, 3.66, 4.22, 0.73, 0.4, '03', {c: TEAL, al: 'c'});
  txt(s, 4.57, 4.22, 0.73, 0.4, '04', {c: TEAL, al: 'c'});
  card(s, 7.32, 4.45, 3.81, 1.54, INDIGO, 0.23);
  txt(s, 7.51, 4.7, 1.98, 0.34, 'Detail Project 01', {c: W});
  body(s, 7.51, 5.08, 3.61, 0.68, L37, W);
}
function slide50(s) {          // Portfolio B07
  photo(s, 6.52, 3.86, 5.26, 2.59);
  photo(s, 9.2, 1.16, 2.58, 2.59);
  txt(s, 1.69, 1.59, 4.83, 1.45, [['Our ', INK], ['Showcase ', TEAL], ['Portfolio', INK]], {sz: 40});
  card(s, 6.53, 1.16, 2.58, 2.59, INDIGO, 0.14);
  txt(s, 6.69, 1.36, 2.29, 1.18, 'AI in Action: Real-World Really Great Applications and Breakthroughs', {sz: 16, c: W});
  oval(s, 8.44, 3.1, 0.47, 0.47, W);
  icon(s, 8.53, 3.2, 0.27, 0.27);
  body(s, 1.69, 4.12, 4.36, 1.27, L42);
  pill(s, 1.77, 5.6, 1.92, 0.47, 'Read More', {sz: 12, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.24});
}
function slide51(s) {          // Device Mockup B01
  photo(s, 6.39, 3.86, 5.18, 2.38, 0.2);
  txt(s, 1.44, 1.49, 5.51, 1.45, [['Creative ', INK], ['Mockup  Section', TEAL]], {sz: 40});
  phone(s, 6.29, 3.71, 5.35, 2.68, true);
  card(s, 7.16, 1.51, 1.88, 2.1, INDIGO, 0.15);
  body(s, 9.69, 2.37, 1.94, 0.97, L9);
  txt(s, 9.69, 1.78, 1.94, 0.44, [['45M', INK], ['+', TEAL]], {sz: 20});
  body(s, 7.27, 2.37, 1.94, 0.97, L9, W);
  body(s, 1.54, 5.19, 4.17, 0.98, L4);
  txt(s, 1.54, 4.35, 1.21, 0.44, [['200', INK], ['K', TEAL]], {sz: 20});
  body(s, 1.54, 4.71, 1.21, 0.37, 'Customers');
  txt(s, 2.89, 4.35, 1.21, 0.44, [['11', INK], ['M', TEAL]], {sz: 20});
  body(s, 2.89, 4.71, 1.21, 0.37, 'Value');
  txt(s, 4.23, 4.35, 1.21, 0.44, [['450', INK], ['+', TEAL]], {sz: 20});
  body(s, 4.23, 4.71, 1.21, 0.37, 'Value');
  oval(s, 7.36, 1.7, 0.47, 0.47, W);
  icon(s, 7.46, 1.8, 0.28, 0.28);
}
function slide52(s) {          // Device Mockup B02
  photo(s, 3.95, 1.13, 2.5, 5.37, 0.22);
  photo(s, 9.68, 4.64, 0.37, 0.37, 0.19);
  photo(s, 9.86, 4.64, 0.37, 0.37, 0.19);
  photo(s, 10.05, 4.64, 0.37, 0.37, 0.19);
  card(s, 1.63, 3.92, 1.93, 1.96, INDIGO, 0.16);
  card(s, 9.54, 4.49, 1.99, 0.67, INDIGO, 0.12);
  txt(s, 7.24, 1.46, 4.68, 1.72, [['Book on Our ', INK], ['App Now!', TEAL]], {sz: 48});
  body(s, 7.24, 5.35, 4.44, 0.68, L18);
  card(s, 7.32, 4.49, 2.12, 0.67, W, 0.12, {shadow: 1});
  txt(s, 7.81, 4.62, 1.57, 0.38, 'Download Now!', {sz: 12, ls: 1.5});
  txt(s, 10.44, 4.64, 1.04, 0.37, '1M+ Users', {sz: 12, c: W, ls: 1.5});
  icon(s, 7.42, 4.62, 0.42, 0.42, TEAL);
  phone(s, 3.79, 1, 2.81, 5.59);
  body(s, 1.76, 2.27, 1.94, 0.97, L9);
  txt(s, 1.76, 1.77, 1.94, 0.44, [['62', INK], ['%', TEAL]], {sz: 20});
  body(s, 1.76, 4.63, 1.94, 0.98, L9, W);
  txt(s, 1.76, 4.13, 1.94, 0.44, '15%', {sz: 20, c: W});
}
function slide53(s) {          // Device Mockup B06
  photo(s, 1.72, 1.03, 2.84, 5.89, 0.47);
  photo(s, 4.33, 1.5, 2.29, 4.93, 0.38);
  phone(s, 1.65, 0.94, 3.04, 6.05);
  txt(s, 7.25, 1.32, 4.49, 1.58, [['Our ', INK], ['Mockup  Section', TEAL]], {sz: 44});
  card(s, 7.25, 5.59, 4.38, 0.59, W, 0.06, {shadow: 1});
  card(s, 7.25, 4.25, 4.38, 0.59, INDIGO, 0.06);
  txt(s, 7.31, 4.36, 0.53, 0.34, '01.', {c: W, b: 1});
  txt(s, 7.79, 4.33, 4.09, 0.38, L23, {sz: 12, c: W, ls: 1.5});
  card(s, 7.25, 4.92, 4.38, 0.59, W, 0.06, {shadow: 1});
  txt(s, 7.31, 5.03, 0.53, 0.34, '02.', {c: TEAL, b: 1});
  txt(s, 7.79, 4.99, 4.09, 0.38, L11, {sz: 12, ls: 1.5});
  txt(s, 7.31, 5.7, 0.53, 0.34, '03.', {c: TEAL, b: 1});
  txt(s, 7.79, 5.66, 4.09, 0.38, L24, {sz: 12, ls: 1.5});
  phone(s, 4.23, 1.46, 2.51, 5.01);
}
function slide54(s) {          // Device Mockup B09
  photo(s, 6.92, 1.04, 2.72, 5.46, 0.34);
  phone(s, 6.84, 0.87, 2.9, 5.75);
  txt(s, 0.39, 0.36, 1.8, 0.3, 'ELEVATE', {sz: 12, c: W, sp: 3});
  txt(s, 1.71, 1.65, 4.95, 1.45, [['Our Great ', INK], ['Mockup Section', TEAL]], {sz: 40});
  txt(s, 9.98, 4.88, 1.77, 0.97, [[L43, GREY, {br: 1}], [L44, GREY]], {sz: 12, f: BODY, ls: 1.5});
  txt(s, 9.98, 4.3, 1.94, 0.5, [['125', INK], ['k', TEAL]], {sz: 24});
  body(s, 1.71, 4.18, 4.36, 1.27, L42);
  pill(s, 1.79, 5.65, 1.92, 0.47, 'Read More', {sz: 12, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.24});
  card(s, 9.87, 1.65, 1.88, 2.1, INDIGO, 0.15);
  body(s, 9.98, 2.51, 1.94, 0.97, L9, W);
  oval(s, 10.07, 1.84, 0.47, 0.47, W);
  icon(s, 10.17, 1.94, 0.28, 0.28);
}
function slide55(s) {          // Device Mockup B10
  photo(s, 6.82, 1.03, 2.6, 5.52, 0.18);
  txt(s, 1.88, 1.74, 3.65, 2.12, [['Creative ', INK], ['Mockups', TEAL], [' & Concepts', INK]], {sz: 40});
  body(s, 1.91, 4.81, 4.18, 0.98, L2);
  phone(s, 6.67, 0.89, 2.9, 5.75);
  card(s, 9.28, 3.7, 2.1, 2.27, INDIGO, 0.15);
  txt(s, 9.41, 5.16, 2.01, 0.57, 'Adapting to an Intelligent World', {c: W});
  oval(s, 10.65, 3.89, 0.56, 0.56, W);
  icon(s, 10.72, 3.96, 0.42, 0.42);
  txt(s, 9.98, 2.21, 1.77, 0.97, [[L43, GREY, {br: 1}], [L44, GREY]], {sz: 12, f: BODY, ls: 1.5});
  txt(s, 9.98, 1.63, 1.94, 0.5, [['82', INK], ['k', TEAL]], {sz: 24});
}
function slide56(s) {          // Contact Us B02
  photo(s, 7.28, 1.35, 4.35, 2.37, 0.19);
  photo(s, 1.79, 3.73, 4.79, 2.49, 0.2);
  txt(s, 1.93, 1.35, 4.35, 1.31, [['Let’s ', INK], ['Connect', TEAL], [' and Explore', INK]], {sz: 36});
  card(s, 6.58, 1.47, 1.56, 1.62, INDIGO, 0.12);
  txt(s, 6.69, 2.26, 1.41, 0.68, [['Great Service ', W, {br: 1}], ['in Town', W]], {sz: 12, ls: 1.5});
  oval(s, 6.76, 1.62, 0.5, 0.5, W);
  txt(s, 6.65, 1.73, 0.73, 0.29, [['24', INK], ['hr', INK, {sz: 9}]], {sz: 10.5, al: 'c'});
  txt(s, 8.74, 4.83, 2.57, 0.42, L26, {f: BODY, ls: 1.5});
  txt(s, 8.74, 5.44, 2.57, 0.42, L27, {f: BODY, ls: 1.5});
  txt(s, 8.74, 4.22, 2.57, 0.42, '+(12) 345 678 910', {f: BODY, ls: 1.5});
  txt(s, 7.59, 4.23, 1.09, 0.42, 'Phone:', {c: TEAL, ls: 1.5});
  txt(s, 7.59, 4.85, 1.09, 0.42, 'Website:', {c: TEAL, ls: 1.5});
  txt(s, 7.59, 5.46, 1.09, 0.42, 'Address:', {c: TEAL, ls: 1.5});
}
function slide57(s) {          // Contact Us B04
  photo(s, 9.03, 1.53, 2.47, 2.22, 0.2);
  txt(s, 1.99, 1.6, 6.93, 1.58, [['Thank ', INK], ['You', TEAL], [' ', INK]], {sz: 88});
  txt(s, 2.02, 3.1, 4.36, 0.4, 'For Your Attention');
  card(s, 7.85, 4.44, 3.65, 1.62, INDIGO, 0.15);
  icon(s, 8.17, 4.76, 0.2, 0.2, W);
  txt(s, 8.68, 4.69, 2.94, 0.3, '(123+) 345 678 910 ', {sz: 12, c: W, sp: 2});
  icon(s, 8.17, 5.14, 0.2, 0.2, W);
  txt(s, 8.69, 5.08, 2.94, 0.3, 'example@email.com', {sz: 12, c: W, sp: 2});
  icon(s, 8.14, 5.55, 0.26, 0.27, W);
  txt(s, 8.68, 5.51, 2.94, 0.3, L45, {sz: 12, c: W, sp: 2});
  body(s, 2.04, 4.61, 4.48, 1.29, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut');
}
function slide58(s) {          // Contact Us B05
  photo(s, 7.97, 1.06, 3.76, 5.38, 0.21);
  txt(s, 1.91, 1.48, 6.79, 1.31, [['Thank ', INK], ['You', TEAL]], {sz: 72});
  card(s, 1.61, 4.36, 6.91, 1.72, INDIGO, 0.17);
  txt(s, 1.91, 4.74, 1.81, 0.34, 'Phone', {c: W, b: 1});
  txt(s, 1.91, 5.05, 1.76, 0.35, '+(12) 345 678 910', {sz: 11, f: BODY, c: W, ls: 1.5});
  txt(s, 1.91, 5.31, 1.76, 0.35, '+(12) 109 876 543 ', {sz: 11, f: BODY, c: W, ls: 1.5});
  txt(s, 6.39, 4.75, 1.81, 0.34, 'Website', {c: W, b: 1});
  txt(s, 6.39, 5.06, 1.93, 0.35, L45, {sz: 11, f: BODY, c: W, ls: 1.5});
  txt(s, 4.09, 4.74, 1.81, 0.34, 'Address', {c: W, b: 1});
  txt(s, 4.09, 5.04, 1.76, 0.63, 'Riverside 09, Malang, Indonesia', {sz: 11, f: BODY, c: W, ls: 1.5});
  body(s, 1.91, 3, 5.52, 0.68, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim');
}
function slide59(s) {          // Contact Us w06
  photo(s, 3.91, 3.94, 7.77, 2.53);
  txt(s, 1.8, 1.45, 4.98, 0.91, [['Get ', INK], ['in Touch!', TEAL]], {sz: 48});
  body(s, 1.8, 2.59, 4.66, 0.68, L31);
  card(s, 1.66, 3.94, 2.18, 2.53, INDIGO, 0.12);
  body(s, 1.85, 4.65, 1.68, 0.97, L8, W);
  txt(s, 1.85, 4.18, 2.05, 0.34, ' Appointment', {c: W});
  pill(s, 1.96, 5.83, 1.29, 0.37, 'Read More', {sz: 11, al: 'c', va: 'm', fill: W, r: 0.19});
  txt(s, 9.09, 2.31, 2.57, 0.42, L26, {f: BODY, ls: 1.5});
  txt(s, 9.09, 2.92, 2.57, 0.42, L27, {f: BODY, ls: 1.5});
  txt(s, 9.09, 1.7, 2.57, 0.42, '+(12) 345 678 910', {f: BODY, ls: 1.5});
  txt(s, 7.94, 1.71, 1.09, 0.42, 'Phone:', {c: TEAL, ls: 1.5});
  txt(s, 7.94, 2.32, 1.09, 0.42, 'Website:', {c: TEAL, ls: 1.5});
  txt(s, 7.94, 2.94, 1.09, 0.42, 'Address:', {c: TEAL, ls: 1.5});
}
function slide60(s) {          // Contact Us B08
  photo(s, 1.91, 3.17, 4.91, 3.15, 0.29);
  txt(s, 2.09, 1.19, 6.88, 1.21, [['Get ', INK], ['in Touch!', TEAL]], {sz: 66});
  txt(s, 8.86, 4.58, 2.57, 0.42, L26, {f: BODY, ls: 1.5});
  txt(s, 8.86, 5.19, 2.57, 0.42, L27, {f: BODY, ls: 1.5});
  txt(s, 8.86, 3.97, 2.57, 0.42, '+(12) 345 678 910', {f: BODY, ls: 1.5});
  txt(s, 7.7, 3.98, 1.09, 0.42, 'Phone:', {c: TEAL, ls: 1.5});
  txt(s, 7.7, 4.6, 1.09, 0.42, 'Website:', {c: TEAL, ls: 1.5});
  txt(s, 7.7, 5.21, 1.09, 0.42, 'Address:', {c: TEAL, ls: 1.5});
  body(s, 9.24, 1.29, 2.39, 0.68, L7);
  pill(s, 9.33, 2.17, 1.92, 0.47, 'Read More', {sz: 12, c: W, al: 'c', va: 'm', fill: INDIGO, r: 0.24});
}/* ---------------------------------------------------------------- deck */
const SLIDES = [
  slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  slide31, slide32, slide33, slide34, slide35, slide36, slide37, slide38, slide39, slide40,
  slide41, slide42, slide43, slide44, slide45, slide46, slide47, slide48, slide49, slide50,
  slide51, slide52, slide53, slide54, slide55, slide56, slide57, slide58, slide59, slide60,
];

/* every slide inherits the master's wordmark and page number */
function chrome(s, n) {
  s.addText('AiTech', {
    x: 0.52, y: 0.36, w: 1.21, h: 0.38,
    fontSize: 12, fontFace: HEAD, color: INK, lineSpacingMultiple: 1.5,
    margin: PAD, isTextBox: true,
  });
  s.addText('Page ' + n, {
    x: 11.85, y: 6.8, w: 0.98, h: 0.29, align: 'right',
    fontSize: 11, fontFace: BODY, color: GREY, margin: PAD, isTextBox: true,
  });
}

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
  pptx.layout = 'DECK';
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.title = 'How AI Technology is Transforming the World';

  SLIDES.forEach(function (fn, i) {
    const s = pptx.addSlide();
    s.background = { color: W };
    chrome(s, i + 1);
    fn(s);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '1624bb61-4823-4849-a970-9fac61e36590_grok_final.pptx'),
  });
}

build().then(function (f) { console.log('wrote', f); },
             function (e) { console.error(e); process.exit(1); });
