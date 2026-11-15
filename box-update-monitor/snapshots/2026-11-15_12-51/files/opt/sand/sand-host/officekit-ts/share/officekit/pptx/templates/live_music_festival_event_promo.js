/**
 * MicroFest — Live Music Event Presentation (30 slides, 16:9 / 13.333in x 7.5in)
 *
 * Rebuilt from scratch with pptxgenjs.  Raster photos in the original deck are
 * replaced by flat "[image]" placeholder rectangles at the same geometry.
 *
 *   node <thisFile>.js   ->   writes the .pptx next to this file
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const PINK      = 'F353AE';   // theme accent1
const PURPLE    = 'B644E4';   // theme accent2
const PINK_LT   = 'F9A1D3';   // theme accent3
const PURPLE_LT = 'D694F0';   // theme accent4
const WHITE     = 'FFFFFF';   // theme accent5 / bg
const GREY      = '7F7F7F';   // theme accent6  - body copy
const GREY60    = 'B2B2B2';   // accent6 lum 60%
const GREY50    = '808080';   // bg1 lum 50%
const INK       = '0F0F0F';   // theme tx1 - headline black
const SUBTLE    = '5E5E5E';   // theme tx2
const PALE      = 'FDDDEF';   // accent1 lum 20% - the big soft-pink panels
const MIST      = 'F2F2F2';
const GRAD_MID  = 'D34CC9';   // flat stand-in for the pink->purple gradient
const PLACEHOLD = 'E8E8E8';   // photo placeholder

const HEAD = 'Raleway SemiBold';   // theme major font
const BODY = 'Roboto';             // theme minor font

// ---------------------------------------------------------------- primitives
const NOLINE = { type: 'none' };

/** linear blend between two hex colours */
function mix(a, b, t) {
  const ch = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  return [0, 1, 2]
    .map(i => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t).toString(16).padStart(2, '0'))
    .join('').toUpperCase();
}

/** flat rectangle */
function flat(s, x, y, w, h, color, transparency) {
  s.addShape('rect', { x, y, w, h, fill: { color, transparency: transparency || 0 }, line: NOLINE });
}

/**
 * The deck's signature pink->purple gradient.  PowerPoint gradients are not
 * exposed by pptxgenjs, so it is painted as a strip of banded rectangles:
 * horizontal bands plus half-transparent vertical bands give the 45deg look.
 */
function gradRect(s, x, y, w, h, opt) {
  const ang = (opt && opt.ang) || 'diag';
  const N = 26;
  for (let i = 0; i < N; i++) {
    s.addShape('rect', {
      x: x + (w * i) / N, y, w: w / N + 0.012, h,
      fill: { color: mix(PINK, PURPLE, (i + 0.5) / N) }, line: NOLINE,
    });
  }
  if (ang === 'diag') {
    for (let i = 0; i < N; i++) {
      s.addShape('rect', {
        x, y: y + (h * i) / N, w, h: h / N + 0.008,
        fill: { color: mix(PINK, PURPLE, (i + 0.5) / N), transparency: 50 }, line: NOLINE,
      });
    }
  }
}

/**
 * Stand-in for a photograph in the original deck (device mock-ups).  Kept
 * semi-transparent so the artwork it sits on still reads through.
 */
function imageBox(s, x, y, w, h) {
  s.addShape('rect', { x, y, w, h, fill: { color: PLACEHOLD, transparency: 35 },
    line: { color: GREY60, width: 0.75, dashType: 'dash' } });
  s.addText('[image]', {
    x, y, w, h, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 11, color: GREY50,
  });
}

// ---------------------------------------------------------------- text
function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: BODY, fontSize: 18, color: INK, valign: 'top' }, o));
}

function rich(s, runs, o) {
  runs.forEach(r => { r.options = Object.assign({ fontFace: BODY }, r.options); });
  s.addText(runs, Object.assign({ fontSize: 18, color: INK, valign: 'top' }, o));
}

/** display title: dark words followed by a pink tail ('\n' starts a new line) */
function heading(s, dark, pink, o) {
  const size = o.fontSize || 32;
  const runs = [];
  const push = (t, color) => {
    t.split('\n').forEach((part, i, all) => {
      runs.push({ text: part, options: { fontFace: HEAD, fontSize: size, color, breakLine: i < all.length - 1 } });
    });
  };
  if (dark) push(dark, INK);
  if (pink) push(pink, PINK);
  s.addText(runs, Object.assign({ valign: 'top' }, o, { fontSize: size }));
}

/** the small pink kicker that sits above almost every title */
function eyebrow(s, x, y, align) {
  s.addText('ENJOY TO LISTEN THE MUSIC', {
    x, y, w: 2.108, h: 0.269, wrap: false, align: align || 'left', valign: 'top',
    fontFace: BODY, fontSize: 10, bold: true, color: PINK,
  });
}

// ---------------------------------------------------------------- deco furniture
/** "Microfest" wordmark plus its short rule, top-left */
function logo(s, x, y) {
  s.addText([
    { text: 'Micro', options: { fontFace: HEAD, fontSize: 9, color: PINK } },
    { text: 'fest',  options: { fontFace: BODY, fontSize: 9, color: INK } },
  ], { x, y, w: 0.766, h: 0.252, wrap: false, valign: 'top' });
  s.addShape('line', { x: x + 0.734, y: y + 0.168, w: 0.276, h: 0, line: { color: PINK, width: 0.75 } });
}

/** "Live Music Now" caption plus the gradient rule, top-right */
function tagline(s) {
  s.addText('Live Music Now', {
    x: 10.21, y: 0.622, w: 1.089, h: 0.252, wrap: false, align: 'right',
    fontFace: BODY, fontSize: 9, color: GREY60,
  });
  gradRect(s, 11.417, 0.761, 1.354, 0.03, { ang: 'h' });
}

/** the four gradient page-progress dots on the right edge */
function dots(s, x, y) {
  for (let i = 0; i < 4; i++) {
    s.addShape('ellipse', {
      x, y: y + i * 0.309, w: 0.084, h: 0.084,
      fill: { color: mix(PINK, PURPLE, i / 3) }, line: NOLINE,
    });
  }
}

/** the three outlined "x" marks, stacked vertically or in a row */
function crosses(s, x, y, size, dir) {
  const step = size * 0.96;
  for (let i = 0; i < 3; i++) {
    s.addShape('mathMultiply', {
      x: x + (dir === 'row' ? i * step : 0),
      y: y + (dir === 'row' ? 0 : i * step),
      w: size, h: size, fill: NOLINE, line: { color: PINK_LT, width: 2.25 },
    });
  }
}

/**
 * The three short rounded "equaliser" bars (one freeform in the original).
 * rot 0 = upright bars, 90 = laid flat, 180 = upright but colour order flipped.
 */
function pills(s, x, y, rot, transparency) {
  const BAR = 0.142, STEP = 0.384, LEN = 0.507;   // one 0.91 x 0.507 freeform in the original
  const flatRot = Math.abs(rot % 180) === 90;
  for (let i = 0; i < 3; i++) {
    const fill = { color: mix(PINK, PURPLE, (rot === 180 ? 2 - i : i) / 2), transparency };
    const o = flatRot
      ? { x: x + 0.202, y: y - 0.202 + i * STEP, w: LEN, h: BAR }   // rotated flat
      : { x: x + i * STEP, y, w: BAR, h: LEN };
    s.addShape('roundRect', Object.assign(o, { rectRadius: BAR / 2, fill, line: NOLINE }));
  }
}

/** the pale circle-over-triangle mark tucked into the bottom-left corner */
function cornerDeco(s, x, y) {
  s.addShape('ellipse',  { x, y, w: 1.174, h: 1.174, fill: { color: PINK, transparency: 80 }, line: NOLINE });
  s.addShape('triangle', { x, y: y + 1.174, w: 1.174, h: 1.174, fill: { color: PINK, transparency: 80 }, line: NOLINE });
}

/** small outlined circle holding a right arrow */
function arrowCircle(s, x, y, d, color) {
  s.addShape('ellipse',    { x, y, w: d, h: d, fill: NOLINE, line: { color, width: 1.25 } });
  s.addShape('rightArrow', { x: x + d * 0.22, y: y + d * 0.36, w: d * 0.56, h: d * 0.28,
    fill: { color }, line: NOLINE });
}

/** white card with the deck's very soft drop shadow */
function card(s, x, y, w, h, color) {
  s.addShape('rect', { x, y, w, h, fill: { color }, line: NOLINE,
    shadow: { type: 'outer', blur: 20, offset: 0, angle: 90, color: '000000', opacity: 0.1 } });
}

/**
 * Small pictogram stand-ins for the vector glyphs of the original deck.
 * Each is assembled from a handful of native shapes inside a size x size box.
 */
function icon(s, kind, x, y, w, h, color, transparency) {
  const F = { color, transparency: transparency || 0 }, u = Math.min(w, h);
  const box = (dx, dy, dw, dh, shape, extra) =>
    s.addShape(shape || 'rect', Object.assign(
      { x: x + dx * w, y: y + dy * h, w: dw * w, h: dh * h, fill: F, line: NOLINE }, extra));
  switch (kind) {
    case 'musicNote':                                    // two beamed quavers
      box(0.02, 0.68, 0.42, 0.30, 'ellipse', { rotate: 340 });
      box(0.56, 0.58, 0.42, 0.30, 'ellipse', { rotate: 340 });
      box(0.36, 0.10, 0.09, 0.73);
      box(0.90, 0.00, 0.09, 0.73);
      box(0.36, 0.02, 0.63, 0.16, 'rect', { rotate: 352 });
      break;
    case 'rockHand':                                     // devil-horns hand
      box(0.14, 0.44, 0.72, 0.56, 'roundRect', { rectRadius: u * 0.16 });
      box(0.02, 0.00, 0.26, 0.60, 'roundRect', { rectRadius: u * 0.13 });
      box(0.72, 0.00, 0.26, 0.60, 'roundRect', { rectRadius: u * 0.13 });
      break;
    case 'equalizer':                                    // stacked level meter
      [0.55, 0.30, 0.95, 0.45].forEach((hh, i) => {
        const cells = Math.round(hh * 7);
        for (let k = 0; k < cells; k++) box(0.03 + i * 0.25, 0.98 - (k + 1) * 0.14, 0.19, 0.09);
      });
      break;
    case 'dancers':                                      // two dancing figures
      box(0.06, 0.28, 0.19, 0.19, 'ellipse');            // left: head, torso, legs
      box(0.10, 0.47, 0.11, 0.24);
      box(0.02, 0.70, 0.11, 0.30, 'rect', { rotate: 20 });
      box(0.16, 0.70, 0.11, 0.30, 'rect', { rotate: 340 });
      box(0.66, 0.10, 0.19, 0.19, 'ellipse');            // right: head, torso, legs
      box(0.70, 0.29, 0.11, 0.24);
      box(0.62, 0.52, 0.11, 0.34, 'rect', { rotate: 20 });
      box(0.78, 0.52, 0.11, 0.34, 'rect', { rotate: 340 });
      box(0.24, 0.40, 0.50, 0.09, 'rect', { rotate: 342 });   // linked arms
      break;
    case 'gear':
      box(0, 0, 1, 1, 'gear9');
      break;
    case 'prevTrack': {                                  // |<< rewind glyph
      box(0.02, 0.16, 0.10, 0.68);
      // rotating a triangle keeps its bounding box, so swap w/h before turning it
      const tw = 0.40 * w, th = 0.68 * h;
      [0.16, 0.55].forEach(dx => s.addShape('triangle', {
        x: x + dx * w + (tw - th) / 2, y: y + 0.16 * h + (th - tw) / 2,
        w: th, h: tw, fill: F, line: NOLINE, rotate: 270 }));
      break;
    }
    case 'personCard':                                   // follower avatar
      box(0.30, 0.02, 0.36, 0.36, 'ellipse');
      box(0.14, 0.44, 0.72, 0.40, 'blockArc', { angleRange: [180, 0], arcThicknessRatio: 1 });
      box(0.00, 0.60, 0.34, 0.34, 'ellipse');
      break;
    case 'chevron':                                      // ">" next arrow
      box(0.10, 0.05, 0.80, 0.90, 'triangle', { rotate: 90 });
      break;
    case 'phone':                                        // tilted handset
      s.addShape('blockArc', { x, y, w, h, fill: F, line: NOLINE,
        angleRange: [300, 120], arcThicknessRatio: 0.4 });
      break;
    case 'envelope':
      s.addShape('rect', { x, y: y + h * 0.15, w, h: h * 0.7,
        fill: { color: WHITE }, line: { color, width: 1.5 } });
      s.addShape('line', { x, y: y + h * 0.15, w: w / 2, h: h * 0.35, line: { color, width: 1.5 } });
      s.addShape('line', { x: x + w / 2, y: y + h * 0.50, w: w / 2, h: -h * 0.35, line: { color, width: 1.5 } });
      break;
    case 'mobile':
      s.addShape('roundRect', { x, y, w, h, rectRadius: u * 0.22, fill: F, line: NOLINE });
      box(0.14, 0.12, 0.72, 0.66, 'rect', { fill: { color: WHITE }, line: NOLINE });
      box(0.40, 0.85, 0.20, 0.09, 'ellipse', { fill: { color: WHITE } });
      break;
    case 'arrowDown':
      box(0.42, 0.00, 0.16, 0.72);
      box(0.05, 0.62, 0.90, 0.38, 'triangle', { rotate: 180 });
      break;
    case 'shapeCluster':                                 // x / triangle / diamond / circle
      box(0.00, 0.00, 0.42, 0.42, 'mathMultiply');
      box(0.56, 0.00, 0.44, 0.42, 'triangle');
      box(0.00, 0.56, 0.42, 0.44, 'diamond');
      box(0.56, 0.56, 0.44, 0.44, 'ellipse');
      break;
  }
}

/** plain horizontal rule */
function rule(s, x, y, w, color, width) {
  s.addShape('line', { x, y, w, h: 0, line: { color, width } });
}

/** a labelled horizontal progress bar ("Type Your Skills Here  ---- 90%") */
function skillBar(s, x, y, label, w, pct) {
  txt(s, label, { x, y, w: 2.348, h: 0.379, align: 'justify', lineSpacingMultiple: 1.5, color: GREY50 });
  flat(s, x + 0.08, y + 0.331, w, 0.192, PURPLE_LT);
  txt(s, pct, { x: x + 0.11 + w, y: y + 0.231, w: 0.6, h: 0.352, lineSpacingMultiple: 1.5, color: GREY50 });
}

// ---------------------------------------------------------------- deck
const pres = new PptxGenJS();
pres.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pres.layout = 'WIDE';
pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
pres.title = 'MicroFest - Live Music Event Presentation';

function slide01(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "31_Title Slide")
  flat(s, 1.625, 0, 3.062, 7.5, PALE);
  gradRect(s, 3.625, 3.75, 3.042, 3.75);
  crosses(s, 5.167, 1.247, 1.042, 'col');
  rich(s, [
    { text: 'Micro', options: { fontSize:72, color:PINK, fontFace:'Raleway Black' } },
    { text: 'Fest', options: { fontSize:72, fontFace:'Raleway Black' } }
  ], { x:7.495, y:2.553, w:5.25, h:1.313, valign:'top' });
  txt(s, 'LIVE MUSIC EVENT PRESENTATION', { x:7.544, y:3.708, w:3.862, h:0.37, valign:'top', wrap:false, fontSize:16, bold:true, color:PINK });
  tagline(s);
  dots(s, 12.665, 5.76);
  arrowCircle(s, 7.62, 4.833, 0.37, GRAD_MID);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit va mus vel euismod leo. Donec commodo et urna', { x:8.094, y:4.705, w:4.231, h:0.625, lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
}

function slide02(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "30_Title Slide")
  flat(s, 1.625, 0, 3.062, 7.5, PALE);
  gradRect(s, 3.604, 3.771, 3.062, 3.062);
  crosses(s, 5.167, 1.247, 1.042, 'col');
  pills(s, 1.17, 5.538, 0, 35);
  tagline(s);
  dots(s, 12.665, 5.76);
  heading(s, 'About Microfest ', 'Live Music Event', { x:7.563, y:1.858, w:4.227, h:1.178 });
  eyebrow(s, 7.575, 1.647);
  txt(s, 'More Information', { x:8.051, y:5.255, w:2.112, h:0.421, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, color:PURPLE });
  arrowCircle(s, 7.646, 5.284, 0.364, PURPLE);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac sem per. Mauris finibus augue id vulputate consectetur adipisci ng elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. sit amet, consectetur adipiscing elit. ', { x:7.563, y:3.358, w:4.231, h:1.459, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
}

function slide03(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "Custom Layout")
  flat(s, 6.667, 0, 5.646, 7.5, PALE);
  gradRect(s, 10.292, 0.75, 3.042, 6.75);
  crosses(s, 7.998, 1.382, 0.778, 'row');
  pills(s, 7.22, 6.153, 90, 35);
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  heading(s, 'About Microfest ', 'Live Music Vision', { x:1.552, y:1.904, w:4.227, h:1.178 });
  eyebrow(s, 1.564, 1.694);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac sem per. Mauris finibus augue id vulputate consectetur', { x:1.573, y:3.12, w:4.231, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'We Are Concern To :', { x:1.568, y:4.386, w:5.159, h:0.421, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:SUBTLE });
  txt(s, 'Lorem ipsum dolor sit amet, cons ectetur adipiscing elit. Vivamus vel euismod ipsum dolor sit', { x:1.585, y:4.888, w:3.943, h:0.625, align:'justify', lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Lorem ipsum dolor sit amet, cons ectetur adipiscing elit. Vivamus vel euismod ipsum dolor sit', { x:1.585, y:5.545, w:3.943, h:0.625, align:'justify', lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:GREY });
}

function slide04(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "1_Custom Layout")
  flat(s, 3.646, 0, 8.021, 7.5, PALE, 65);
  gradRect(s, 10.646, 0, 2.688, 7.5);
  crosses(s, 0.477, 1.361, 0.778, 'row');
  pills(s, 4.212, 5.153, 90, 35);
  icon(s, 'arrowDown', 12.59, 6.387, 0.183, 0.376, WHITE);
  rich(s, [
    { text: 'Microfest', options: { fontSize:20, color:WHITE, fontFace:HEAD, breakLine:true } },
    { text: 'Live Music', options: { fontSize:20, color:WHITE, fontFace:HEAD } }
  ], { x:10.646, y:3.992, w:2.173, h:0.774, align:'right', valign:'top' });
  txt(s, 'Page 4', { x:11.783, y:2.942, w:1.035, h:0.303, align:'right', valign:'top', fontSize:12, color:WHITE, fontFace:HEAD });
  icon(s, 'shapeCluster', 12.058, 0.681, 0.623, 0.62, WHITE, 45);
  heading(s, 'About Microfest ', 'Live Music Mission', { x:5.53, y:1.957, w:4.227, h:1.178 });
  eyebrow(s, 5.542, 1.746);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac sem per. Mauris finibus augue id vulputate consectetur dolor sit amet, consectetur adipiscing elit. Vivamus vel', { x:5.551, y:3.173, w:4.231, h:1.181, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Lorem ipsum dolor sitam et, consectetur adipi', { x:5.564, y:5.484, w:1.961, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  rich(s, [
    { text: '792', options: { fontSize:24, bold:true, color:PINK } },
    { text: '+', options: { fontSize:24, bold:true, color:PINK, superscript:true } }
  ], { x:5.564, y:4.924, w:1.709, h:0.64, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, 'Lorem ipsum dolor sitam et, consectetur adipi ', { x:7.786, y:5.484, w:1.961, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  rich(s, [
    { text: '+', options: { fontSize:24, bold:true, color:PURPLE, superscript:true } },
    { text: '0,141%', options: { fontSize:24, bold:true, color:PURPLE } }
  ], { x:7.786, y:4.924, w:1.709, h:0.64, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, 'Aspects 01', { x:5.564, y:4.772, w:1, h:0.348, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, italic:true, color:GREY });
  txt(s, 'Aspects 02', { x:7.786, y:4.772, w:1, h:0.348, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, italic:true, color:GREY });
}

function slide05(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "2_Custom Layout")
  flat(s, 0, 3.811, 13.333, 3.75, PALE);
  gradRect(s, 0.646, 0.75, 5.104, 6);
  crosses(s, 6.667, 0.361, 0.778, 'row');
  pills(s, 12.233, 3.497, 90, 35);
  txt(s, 'About Microfest Live Music History', { x:1.267, y:1.904, w:4.227, h:1.178, valign:'top', fontSize:32, color:WHITE, fontFace:HEAD });
  eyebrow(s, 1.267, 1.694);
  txt(s, '90%', { x:6.752, y:4.999, w:0.94, h:0.729, lineSpacingMultiple:1.5, valign:'top', fontSize:28, bold:true, color:PURPLE });
  s.addShape('upArrow', { x:6.667, y:5.388, w:0.099, h:0.151, fill:{color:PINK}, line:{type:'none'} });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac sem per. Mauris finibus augue id vulputate consecteturconsectetur adipiscing elit. Vivamus vel eu', { x:6.581, y:5.703, w:5.288, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, '2021 -', { x:1.267, y:3.369, w:1.283, h:0.467, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:16, bold:true, color:WHITE });
  txt(s, 'Lorem ipsum dolor sit amet, conse ctetur adipiscing elit. Vivamus vel euismod leo. Donec commod', { x:1.267, y:3.836, w:3.823, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  txt(s, '2022 -', { x:1.267, y:4.795, w:1.283, h:0.467, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:16, bold:true, color:WHITE });
  txt(s, 'Lorem ipsum dolor sit amet, conse ctetur adipiscing elit. Vivamus vel euismod leo. Donec commod', { x:1.267, y:5.261, w:3.823, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  txt(s, 'Writes Simple Description ', { x:7.669, y:5.186, w:1.537, h:0.505, valign:'top', fontSize:12, bold:true, color:SUBTLE });
}

function slide06(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "3_Custom Layout")
  flat(s, 0, 0, 3.684, 7.5, PALE);
  gradRect(s, 2.688, 4.083, 3.979, 2.667);
  crosses(s, 5.125, 0.75, 1.042, 'col');
  pills(s, 1.128, 5.455, 0, 35);
  tagline(s);
  dots(s, 12.665, 5.76);
  txt(s, 'What We Provide For ?', { x:8.029, y:4.297, w:4.675, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:SUBTLE });
  txt(s, 'Check Our Catalog Website', { x:8.029, y:5.416, w:4.675, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:SUBTLE });
  heading(s, 'This What We\n', 'Provide To Customer', { x:7.519, y:1.869, w:4.675, h:1.178 });
  eyebrow(s, 7.531, 1.659);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac sem per. Mauris finibus augue id vulputate consectetur', { x:7.54, y:3.085, w:4.231, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Lorem ipsum dolor sit amet, consect etur adipi scing elit. Vivamus vel euismod leo. ', { x:8.029, y:4.645, w:3.635, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, '01', { x:7.499, y:4.26, w:0.612, h:0.55, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:20, bold:true, color:PINK });
  txt(s, 'Lorem ipsum dolor sit amet, consect etur adipis cing elit. Vivamus vel euismod leo. ', { x:8.029, y:5.763, w:3.635, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, '02', { x:7.499, y:5.379, w:0.654, h:0.555, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:20, bold:true, color:PINK });
}

function slide07(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "4_Custom Layout")
  flat(s, 9.649, 0, 3.684, 7.5, PALE);
  logo(s, 0.563, 0.577);
  gradRect(s, 7.688, 1.694, 4, 4.047);
  crosses(s, 4.75, 3.354, 0.778, 'row');
  pills(s, 0.238, 6.111, 90, 35);
  heading(s, 'About Microfest ', 'Key Milestones', { x:1.552, y:1.904, w:4.227, h:1.178 });
  eyebrow(s, 1.564, 1.694);
  txt(s, 'Description Here', { x:8.559, y:2.326, w:2.94, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:WHITE });
  txt(s, 'Description Here', { x:8.559, y:3.941, w:2.94, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:WHITE });
  txt(s, 'Lorem ipsum dolor sit amet, consec tetur adipiscing elit. Vivamus vel eu ismod leo. Donec commodo.', { x:8.559, y:2.665, w:2.653, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  txt(s, 'Lorem ipsum dolor sit amet, consec tetur adipiscing elit. Vivamus vel eu ismod leo. Donec commodo.', { x:8.559, y:4.269, w:2.653, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  icon(s, 'musicNote', 8.174, 2.403, 0.347, 0.395, WHITE);
  icon(s, 'rockHand', 8.228, 4.089, 0.219, 0.36, WHITE);
}

function slide08(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "5_Custom Layout")
  gradRect(s, 0.667, 0, 2.979, 4.75);
  crosses(s, 5.146, 1.875, 1.042, 'col');
  pills(s, 1.212, 6.038, 90, 35);
  tagline(s);
  dots(s, 12.665, 5.76);
  heading(s, 'Let’s Know More About\n', 'Microfest Company', { x:6.667, y:1.831, w:5.914, h:1.178 });
  eyebrow(s, 6.679, 1.621);
  rich(s, [
    { text: '12.35 K -', options: { fontSize:16, bold:true, color:PINK } },
    { text: ' ', options: { fontSize:11, bold:true, color:PINK } },
    { text: 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac ipsum dolor ipsum dolor sit amet, consectetur adi vel euismod leo', options: { fontSize:11, color:GREY } }
  ], { x:6.667, y:3.352, w:5, h:1.029, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adi pis cing elit. Vivamus vel euismod leo. Donec com modo et urna ac', { x:6.679, y:4.495, w:5, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Show Details', { x:6.748, y:5.593, w:1.497, h:0.286, align:'center', valign:'top', line:{color:PURPLE,width:1}, fontSize:11, bold:true, color:PURPLE });
  txt(s, 'Next Page', { x:8.296, y:5.593, w:1.497, h:0.286, align:'center', valign:'top', fill:{color:PURPLE}, line:{color:PURPLE,width:1}, fontSize:11, bold:true, color:WHITE });
}

function slide09(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "6_Custom Layout")
  gradRect(s, 8.604, 1.771, 3.042, 5.729);
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  crosses(s, 7.257, 1.06, 0.778, 'col');
  pills(s, 12.257, 5.506, 90, 35);
  heading(s, 'About Microfest ', 'Service Types ', { x:1.567, y:2.039, w:4.227, h:1.178 });
  eyebrow(s, 1.579, 1.829);
  rich(s, [
    { text: 'First Service Here, ', options: { fontSize:14, bold:true, color:SUBTLE } },
    { text: 'Lorem ipsum dolor sit amet, conse ctetur adipi scing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate semper. Mauris finibus ', options: { fontSize:11, color:GREY } }
  ], { x:1.567, y:3.496, w:5.266, h:0.979, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  rich(s, [
    { text: 'Second Service Here, ', options: { fontSize:14, bold:true, color:SUBTLE } },
    { text: 'Lorem ipsum dolor sit amet, conse ctetur adip iscing elit. Vivamus vel euismod leo. Donec commodo et urna ac semper. Mauris finibus augue id vulputate semper. Mauris finibus ', options: { fontSize:11, color:GREY } }
  ], { x:1.579, y:4.692, w:5.266, h:0.979, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
}

function slide10(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "7_Custom Layout")
  gradRect(s, 0, 0, 4.684, 7.5);
  tagline(s);
  dots(s, 12.665, 5.76);
  crosses(s, 5.82, 2.623, 0.735, 'col');
  heading(s, 'About Microfest ', 'Service Details', { x:7.257, y:1.649, w:4.227, h:1.178 });
  eyebrow(s, 7.27, 1.439);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac. Donec commodo et urna ac semper. semper. ', { x:7.877, y:3.542, w:4.006, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  rich(s, [
    { text: '1', options: { fontSize:14, bold:true, color:SUBTLE } },
    { text: 'ST', options: { fontSize:14, bold:true, color:SUBTLE, superscript:true } },
    { text: ' Recommendation ', options: { fontSize:14, bold:true, color:SUBTLE } }
  ], { x:7.877, y:3.103, w:4.006, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  icon(s, 'dancers', 7.257, 4.719, 0.578, 0.578, PURPLE);
  icon(s, 'equalizer', 7.27, 3.185, 0.493, 0.493, PURPLE);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac. Donec commodo et urna ac semper. semper. ', { x:7.877, y:5.158, w:4.006, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  rich(s, [
    { text: '2', options: { fontSize:14, bold:true, color:SUBTLE } },
    { text: 'nd', options: { fontSize:14, bold:true, color:SUBTLE, superscript:true } },
    { text: ' Best Service', options: { fontSize:14, bold:true, color:SUBTLE } }
  ], { x:7.877, y:4.719, w:4.006, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
}

function slide11(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "8_Custom Layout")
  gradRect(s, 7.667, 3.75, 4.266, 3.75);
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  crosses(s, 8.257, 0.24, 0.778, 'col');
  card(s, 5.616, 4.495, 4.067, 2.182, WHITE);
  card(s, 1.33, 4.495, 4.067, 2.182, WHITE);
  s.addShape('rect', { x:1.33, y:4.936, w:0.05, h:1.3, fill:{color:PINK}, line:{type:'none'} });
  heading(s, 'About Microfest ', 'Advangates Services', { x:1.573, y:1.858, w:5.094, h:1.178 });
  eyebrow(s, 1.585, 1.647);
  txt(s, 'Lorem Ipsum is simply tex simply dummy text of the printing and typese tting industry. Dum my text of the printing and typese', { x:1.683, y:5.559, w:3.337, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'About Detail Here', { x:1.683, y:5.187, w:2.328, h:0.372, lineSpacingMultiple:1.5, valign:'top', fontSize:12, bold:true, color:'404040' });
  txt(s, '26 M', { x:3.431, y:4.71, w:1.613, h:0.729, align:'right', lineSpacingMultiple:1.5, valign:'top', fontSize:28, bold:true, color:PURPLE });
  txt(s, 'Lorem Ipsum is simply tex simply dummy text of the printing and typese tting industry. Dum my text of the printing and typese', { x:5.968, y:5.559, w:3.337, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'About Detail Here', { x:5.968, y:5.187, w:2.328, h:0.372, lineSpacingMultiple:1.5, valign:'top', fontSize:12, bold:true, color:'404040' });
  txt(s, '18%', { x:8.094, y:4.71, w:1.236, h:0.729, align:'right', lineSpacingMultiple:1.5, valign:'top', fontSize:28, bold:true, color:PURPLE });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac. Donec commodo et urna ac semper. semper. euismod leo. Donec commodo et urna ac. ', { x:1.573, y:3.181, w:5.538, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
}

function slide12(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "9_Custom Layout")
  gradRect(s, 0, 0, 3.688, 7.5);
  card(s, 1.583, 3.75, 3.094, 2.979, WHITE);
  s.addShape('ellipse', { x:2.677, y:4.197, w:0.906, h:0.906, fill:{color:PINK,transparency:45}, line:{type:'none'} });
  icon(s, 'gear', 2.967, 4.487, 0.326, 0.326, WHITE);
  crosses(s, 5.24, 1.271, 0.849, 'col');
  tagline(s);
  dots(s, 12.665, 5.76);
  heading(s, 'Knowing Our\n', 'Recommended', { x:6.958, y:1.858, w:4.227, h:1.178 });
  eyebrow(s, 6.971, 1.647);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac sem per. Mauris fini bus augue id vulputate consectetur adipisci ng elit. ', { x:6.954, y:3.929, w:4.796, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  icon(s, 'prevTrack', 6.955, 3.348, 0.528, 0.528, PURPLE);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac sem per. ', { x:6.971, y:4.924, w:4.796, h:0.625, align:'justify', lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. Donec commodo et urna ac sem per. ', { x:6.954, y:5.642, w:4.796, h:0.625, align:'justify', lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Most Used Services', { x:7.549, y:3.309, w:2.112, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, color:PURPLE });
  txt(s, 'Lorem ipsum dolor sit', { x:7.549, y:3.566, w:4.796, h:0.348, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, italic:true, color:GREY });
  txt(s, 'Lorem ipsum dolor sit amet, con sectetur adipiscing elit. Vivamus amet, consectetur adipiscing con', { x:1.714, y:5.548, w:2.833, h:0.903, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Insert Title – 76,2 %', { x:1.937, y:5.219, w:2.385, h:0.37, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:12, bold:true, color:SUBTLE });
}

function slide13(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "10_Custom Layout")
  gradRect(s, 1.667, 5.206, 3, 1.544);
  gradRect(s, 5.167, 5.206, 3, 1.544);
  gradRect(s, 8.667, 5.206, 3, 1.544);
  crosses(s, 11.329, 2.333, 0.675, 'col');
  pills(s, 1.212, 3.841, 90, 35);
  heading(s, 'Meet Our ', 'Team', { x:4.553, y:0.92, w:4.227, h:0.64, align:'center' });
  eyebrow(s, 5.613, 0.71, 'center');
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus', { x:5.44, y:5.911, w:2.453, h:0.625, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  rich(s, [
    { text: 'Veronica ', options: { fontSize:16, bold:true, color:WHITE } },
    { text: '/ ', options: { fontSize:14, color:WHITE } },
    { text: 'Stage Manager ', options: { fontSize:11, color:WHITE } }
  ], { x:5.44, y:5.419, w:2.453, h:0.46, align:'center', lineSpacingMultiple:1.5, valign:'top' });
  rule(s, 5.613, 5.902, 2.108, WHITE, 1);
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus', { x:8.94, y:5.911, w:2.453, h:0.625, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  rich(s, [
    { text: 'Agustinous ', options: { fontSize:16, bold:true, color:WHITE } },
    { text: '/ ', options: { fontSize:14, color:WHITE } },
    { text: 'Soundman', options: { fontSize:11, color:WHITE } }
  ], { x:8.94, y:5.419, w:2.453, h:0.46, align:'center', lineSpacingMultiple:1.5, valign:'top' });
  rule(s, 9.135, 5.902, 2.057, WHITE, 1);
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus', { x:1.94, y:5.911, w:2.453, h:0.625, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  rich(s, [
    { text: 'Laura Ann ', options: { fontSize:16, bold:true, color:WHITE } },
    { text: '/ ', options: { fontSize:14, color:WHITE } },
    { text: 'CEO', options: { fontSize:11, color:WHITE } }
  ], { x:1.94, y:5.419, w:2.453, h:0.46, align:'center', lineSpacingMultiple:1.5, valign:'top' });
  rule(s, 2.375, 5.902, 1.567, WHITE, 1);
}

function slide14(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "11_Custom Layout")
  flat(s, 0, 0, 4.667, 7.5, PALE);
  gradRect(s, 0.708, 0.724, 2.979, 2.069);
  tagline(s);
  crosses(s, 5.329, 2.562, 0.675, 'col');
  pills(s, 1.212, 5.062, 90, 35);
  dots(s, 12.665, 5.76);
  heading(s, 'About Microfest ', 'Live Music CEO', { x:6.667, y:1.681, w:4.227, h:1.178 });
  eyebrow(s, 6.679, 1.471);
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam vulputate dignissim. Enimagi cinec dui nunc praesent tristique magna sit amet. Risus nunc praesent tristique magna sit amet. Risus', { x:6.667, y:3.488, w:5.125, h:0.934, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Laura Ann', { x:6.667, y:3.151, w:1.233, h:0.37, valign:'top', wrap:false, fontSize:16, bold:true, color:PINK_LT });
  s.addShape('rect', { x:6.746, y:5.09, w:4.233, h:0.192, fill:{color:PURPLE_LT}, line:{type:'none'} });
  s.addShape('rect', { x:6.746, y:5.785, w:4.4, h:0.192, fill:{color:PURPLE_LT}, line:{type:'none'} });
  txt(s, 'Type Your Skills Here', { x:6.667, y:4.759, w:2.348, h:0.379, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, 'Type Your Experience Here', { x:6.679, y:5.433, w:2.348, h:0.379, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, '90%', { x:10.996, y:4.99, w:0.544, h:0.352, lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, '95%', { x:11.189, y:5.677, w:0.603, h:0.352, lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
}

function slide15(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "12_Custom Layout")
  gradRect(s, 9.667, 3.75, 3.667, 3.75);
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  crosses(s, 7.329, 1.663, 0.675, 'col');
  pills(s, 11.774, 0.491, 180, 35);
  heading(s, 'Meet Our Famous\n', 'Stage Manager Event', { x:1.701, y:1.681, w:4.965, h:1.178 });
  eyebrow(s, 1.714, 1.471);
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam vulputate dignissim. Enimagi cinec dui nunc praesent tristique magna sit amet. Risusmet. Risus pretiu mquam ', { x:1.701, y:3.488, w:4.965, h:0.934, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Veronica', { x:1.701, y:3.151, w:1.103, h:0.37, valign:'top', wrap:false, fontSize:16, bold:true, color:PINK_LT });
  s.addShape('rect', { x:1.781, y:5.09, w:4.052, h:0.192, fill:{color:PURPLE_LT}, line:{type:'none'} });
  s.addShape('rect', { x:1.781, y:5.785, w:4.318, h:0.192, fill:{color:PURPLE_LT}, line:{type:'none'} });
  txt(s, 'Type Your Skills Here', { x:1.701, y:4.759, w:2.348, h:0.379, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, 'Type Your Experience Here', { x:1.713, y:5.433, w:2.348, h:0.379, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, '90%', { x:5.851, y:4.99, w:0.544, h:0.352, lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, '99%', { x:6.134, y:5.677, w:0.603, h:0.352, lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
}

function slide16(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "13_Custom Layout")
  gradRect(s, 2.688, 0.74, 3, 4.01);
  crosses(s, 4.35, 4.932, 0.675, 'col');
  pills(s, 1.233, 1.476, 180, 35);
  tagline(s);
  dots(s, 12.665, 5.76);
  heading(s, 'The Last, Our Best\n', 'Soundman Here!', { x:6.524, y:1.681, w:4.227, h:1.178 });
  eyebrow(s, 6.537, 1.471);
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam vulputate dignissim. Enimagi cinec dui nunc praesent tristique magna sit amet. Risus tristique magna sit amet. Risus', { x:6.524, y:3.488, w:5.272, h:0.934, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Agustinous', { x:6.524, y:3.151, w:1.352, h:0.37, valign:'top', wrap:false, fontSize:16, bold:true, color:PINK_LT });
  s.addShape('rect', { x:6.604, y:5.09, w:4, h:0.192, fill:{color:PURPLE_LT}, line:{type:'none'} });
  s.addShape('rect', { x:6.604, y:5.785, w:4.253, h:0.192, fill:{color:PURPLE_LT}, line:{type:'none'} });
  txt(s, 'Type Your Skills Here', { x:6.524, y:4.759, w:2.348, h:0.379, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, 'Type Your Experience Here', { x:6.536, y:5.433, w:2.348, h:0.379, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, '83%', { x:10.644, y:4.99, w:0.544, h:0.352, lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, '92%', { x:10.902, y:5.677, w:0.603, h:0.352, lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
}

function slide17(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "29_Title Slide")
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  tagline(s);
  dots(s, 12.665, 5.76);
  heading(s, 'Our Pricing ', 'Table', { x:4.553, y:1.441, w:4.227, h:0.64, align:'center' });
  eyebrow(s, 5.613, 1.231, 'center');
  card(s, 1.982, 2.638, 2.737, 3.847, PINK);
  card(s, 8.614, 2.638, 2.737, 3.847, PINK);
  card(s, 5.298, 2.638, 2.737, 3.847, PURPLE);
  txt(s, 'Enim nec dui nunc', { x:2.189, y:4.681, w:1.924, h:0.352, lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:WHITE });
  txt(s, 'Magna sit amet risus pro', { x:2.189, y:4.979, w:2.359, h:0.348, lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:WHITE });
  txt(s, 'Risus pretium quam', { x:2.189, y:5.278, w:2.201, h:0.348, lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:WHITE });
  rich(s, [
    { text: '$30', options: { fontSize:44, bold:true, color:WHITE } },
    { text: '/month', options: { fontSize:11, bold:true, color:WHITE } }
  ], { x:2.25, y:3.583, w:2.201, h:1.107, align:'center', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, 'Enim nec dui nunc', { x:5.505, y:4.681, w:1.924, h:0.349, lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:WHITE });
  txt(s, 'Magna sit amet risus pro', { x:5.505, y:4.979, w:2.359, h:0.349, lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:WHITE });
  txt(s, 'Risus pretium quam', { x:5.505, y:5.278, w:2.201, h:0.349, lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:WHITE });
  rich(s, [
    { text: '$39', options: { fontSize:44, bold:true, color:WHITE } },
    { text: '/ month', options: { fontSize:11, bold:true, color:WHITE } }
  ], { x:5.566, y:3.583, w:2.201, h:1.107, align:'center', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, 'Enim nec dui nunc', { x:8.82, y:4.681, w:1.924, h:0.349, lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:WHITE });
  txt(s, 'Magna sit amet risus pro', { x:8.82, y:4.979, w:2.359, h:0.349, lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:WHITE });
  txt(s, 'Risus pretium quam', { x:8.82, y:5.278, w:2.201, h:0.349, lineSpacingMultiple:1.5, bullet:{characterCode:'2713',indent:13.5}, valign:'top', fontSize:11, color:WHITE });
  rich(s, [
    { text: '$40', options: { fontSize:44, bold:true, color:WHITE } },
    { text: '/ month', options: { fontSize:11, bold:true, color:WHITE } }
  ], { x:8.882, y:3.583, w:2.201, h:1.107, align:'center', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, 'Purchase Now', { x:2.687, y:5.834, w:1.328, h:0.419, align:'center', valign:'middle', line:{color:'E5E5E5',width:1}, fontSize:7, bold:true, color:WHITE });
  txt(s, 'Purchase Now', { x:6.003, y:5.834, w:1.328, h:0.419, align:'center', valign:'middle', line:{color:'E5E5E5',width:1}, fontSize:7, bold:true, color:WHITE });
  txt(s, 'Purchase Now', { x:9.318, y:5.834, w:1.328, h:0.419, align:'center', valign:'middle', line:{color:'E5E5E5',width:1}, fontSize:7, bold:true, color:WHITE });
  txt(s, 'The Pricing One', { x:2.34, y:3.059, w:2.022, h:0.404, align:'center', valign:'top', wrap:false, fontSize:18, bold:true, color:WHITE });
  txt(s, 'Updated At 2022', { x:2.589, y:3.384, w:1.524, h:0.307, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:9, italic:true, color:WHITE });
  txt(s, 'The Pricing Two', { x:5.64, y:3.059, w:2.053, h:0.404, align:'center', valign:'top', wrap:false, fontSize:18, bold:true, color:WHITE });
  txt(s, 'Updated At 2022', { x:5.905, y:3.384, w:1.524, h:0.307, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:9, italic:true, color:WHITE });
  txt(s, 'The Pricing Three', { x:8.866, y:3.059, w:2.234, h:0.404, align:'center', valign:'top', wrap:false, fontSize:18, bold:true, color:WHITE });
  txt(s, 'Updated At 2022', { x:9.221, y:3.384, w:1.524, h:0.307, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:9, italic:true, color:WHITE });
}

function slide18(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "15_Custom Layout")
  flat(s, 8.729, 0, 4, 7.5, PALE);
  gradRect(s, 9.667, 0.74, 3.667, 3.01);
  crosses(s, 7.338, 4.196, 0.675, 'col');
  pills(s, 8.288, 1.476, 180, 35);
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  heading(s, 'About Microfest ', 'Product Service', { x:1.535, y:1.737, w:4.227, h:1.178 });
  eyebrow(s, 1.547, 1.526);
  txt(s, '01', { x:1.534, y:2.915, w:1.181, h:1.178, lineSpacingMultiple:1.5, valign:'top', fontSize:48, bold:true, color:MIST });
  rich(s, [
    { text: 'Product 01,', options: { fontSize:11, bold:true, color:PURPLE } },
    { text: ' ', options: { fontSize:11, color:PURPLE } },
    { text: 'Lorem ipsum do lor sit amet, consectetur adipi scing elit. Vivamus', options: { fontSize:11, color:GREY } }
  ], { x:1.975, y:3.495, w:2.229, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, '02', { x:4.296, y:2.915, w:1.181, h:1.178, lineSpacingMultiple:1.5, valign:'top', fontSize:48, bold:true, color:MIST });
  rich(s, [
    { text: 'Product 02', options: { fontSize:11, bold:true, color:PURPLE } },
    { text: ', Lorem ipsum do lor sit amet, consectetur adipi scing elit. Vivamus', options: { fontSize:11, color:GREY } }
  ], { x:4.737, y:3.495, w:2.229, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, '03', { x:1.534, y:4.435, w:1.181, h:1.178, lineSpacingMultiple:1.5, valign:'top', fontSize:48, bold:true, color:MIST });
  rich(s, [
    { text: 'Product 03', options: { fontSize:11, bold:true, color:PURPLE } },
    { text: '.', options: { fontSize:11, bold:true, color:GREY } },
    { text: ' Lorem ipsum do lor sit amet, consectetur adipi scing elit. Vivamus', options: { fontSize:11, color:GREY } }
  ], { x:1.975, y:5.015, w:2.229, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, '04', { x:4.296, y:4.435, w:1.181, h:1.178, lineSpacingMultiple:1.5, valign:'top', fontSize:48, bold:true, color:MIST });
  rich(s, [
    { text: 'Product 04, ', options: { fontSize:11, bold:true, color:PURPLE } },
    { text: 'Lorem ipsum do lor sit amet, consectetur adipi scing elit. Vivamus', options: { fontSize:11, color:GREY } }
  ], { x:4.737, y:5.015, w:2.229, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
}

function slide19(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "16_Custom Layout")
  gradRect(s, 0, 3.719, 10.646, 3.01);
  crosses(s, 0.341, 1.422, 0.694, 'col');
  pills(s, 11.736, 0.517, 180, 35);
  heading(s, 'Keep Microfest\n', 'Innovation Event', { x:5.32, y:1.399, w:4.227, h:1.178 });
  eyebrow(s, 5.32, 1.189);
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam vulputate dignissim. Enimagi cinec dui nunc praesent tristique magna sit amet. Risus tristique magna sit amet. Risus', { x:5.32, y:4.698, w:4.656, h:0.934, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  txt(s, '+76.921 %', { x:5.32, y:4.193, w:2.496, h:0.505, lineSpacingMultiple:1.5, valign:'top', fontSize:18, bold:true, color:WHITE });
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam vulputate dignissim. Enimagi cinec dui nunc', { x:5.32, y:2.681, w:4.656, h:0.625, lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
}

function slide20(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "17_Custom Layout")
  flat(s, 0, 4.738, 13.333, 2.762, PALE);
  gradRect(s, 6.774, 2.75, 2.85, 4.083);
  logo(s, 0.563, 0.577);
  tagline(s);
  txt(s, 'Lorem ipsum dolor sit amet, con sectetur adipiscing elit. Vivam us vel euismod leo. ', { x:6.959, y:3.764, w:2.48, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  rich(s, [
    { text: '01. ', options: { fontSize:18, bold:true, color:WHITE } },
    { text: 'Subtitle Here', options: { fontSize:12, bold:true, color:WHITE } },
    { text: ' ', options: { fontSize:18, bold:true, color:WHITE } }
  ], { x:6.959, y:3.28, w:2.202, h:0.505, lineSpacingMultiple:1.5, valign:'top' });
  txt(s, 'Lorem ipsum dolor sit amet, con sectetur adipiscing elit. Vivam us vel euismod leo. ', { x:6.959, y:5.4, w:2.48, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  rich(s, [
    { text: '02. ', options: { fontSize:18, bold:true, color:WHITE } },
    { text: 'Subtitle Here', options: { fontSize:12, bold:true, color:WHITE } },
    { text: ' ', options: { fontSize:18, bold:true, color:WHITE } }
  ], { x:6.959, y:4.916, w:2.202, h:0.505, lineSpacingMultiple:1.5, valign:'top' });
  heading(s, 'Launching Our ', 'New Product', { x:2.938, y:1.598, w:7.458, h:0.64, align:'center' });
  eyebrow(s, 5.602, 1.387, 'center');
}

function slide21(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "18_Custom Layout")
  flat(s, 1.646, 0, 4.041, 7.5, PALE);
  gradRect(s, 0, 0.761, 4.708, 3.003);
  tagline(s);
  dots(s, 12.665, 5.76);
  heading(s, 'About Planning & ', 'Discuss Live Music', { x:7.562, y:1.916, w:5.208, h:1.178 });
  eyebrow(s, 7.575, 1.706);
  rich(s, [
    { text: 'Description Here', options: { fontSize:14, bold:true, color:PURPLE, breakLine:true } },
    { text: 'Aliquaet nibh praesent tristique ma gna sit amet. Risus sit', options: { fontSize:11, color:GREY } }
  ], { x:9.068, y:3.386, w:2.619, h:1.01, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  rich(s, [
    { text: 'Aliquaet nibh praesent tristique magna sit amet. Risus pretium quam vulputate dignissim', options: { fontSize:11, color:GREY } },
    { text: '. Enimagi cinec dui nunc ', options: { fontSize:11, bold:true, color:GREY } },
    { text: 'mattis enim ut tellus maelem entum iena nec dui ultricestristique magna sit amet. ', options: { fontSize:11, color:GREY } }
  ], { x:7.563, y:4.613, w:4.227, h:1.181, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, '81%', { x:7.575, y:3.109, w:1.566, h:1.178, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:48, bold:true, color:PURPLE });
}

function slide22(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "19_Custom Layout")
  flat(s, 6.667, 0, 4.041, 7.5, PALE);
  gradRect(s, 7.646, 0.729, 5.021, 4.021);
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  crosses(s, 8.329, 5.125, 0.675, 'col');
  heading(s, 'Some Reviews\n', 'In Our Catalog', { x:1.623, y:1.925, w:4.227, h:1.178 });
  eyebrow(s, 1.635, 1.715);
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretium quam vulputate dignissim. Enimagi cinec dui nunc mattis enim ut tellus maelem entum iena nec dui ultricestristique', { x:1.623, y:3.877, w:4.499, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Description Here', { x:1.623, y:3.485, w:2.056, h:0.337, valign:'top', fontSize:14, bold:true, color:SUBTLE });
  rich(s, [
    { text: 'Aliquaet nibh praesent tristique ', options: { fontSize:11, italic:true, color:GREY } },
    { text: 'magna sit amet. Risus pre tium', options: { fontSize:14, italic:true, color:PURPLE } },
    { text: ' ', options: { fontSize:14, italic:true, color:PINK } },
    { text: 'quam vulputate dignissim. Enimagi cinec dui nunc mattis', options: { fontSize:11, italic:true, color:GREY } }
  ], { x:1.574, y:5.017, w:4.499, h:0.769, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, '+124.27 K Reviewers', { x:8.667, y:2.06, w:4.499, h:0.46, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:16, italic:true, color:WHITE });
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretium quam vulputate dignissim. ', { x:8.667, y:1.461, w:3.615, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
}

function slide23(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "20_Custom Layout")
  gradRect(s, 7.688, 0.688, 5, 6.812);
  logo(s, 0.563, 0.577);
  crosses(s, 0.309, 4.542, 0.675, 'col');
  heading(s, 'About Microfest ', 'Insight Social Media', { x:1.573, y:1.925, w:4.575, h:1.178 });
  eyebrow(s, 1.585, 1.715);
  icon(s, 'personCard', 8.248, 2.614, 0.355, 0.355, WHITE);
  icon(s, 'rockHand', 8.248, 1.403, 0.355, 0.355, WHITE);
  rich(s, [
    { text: '+51.23 Followers Youtube', options: { fontSize:12, bold:true, color:WHITE, breakLine:true } },
    { text: 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam', options: { fontSize:11, color:WHITE } }
  ], { x:8.754, y:1.298, w:2.898, h:0.928, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  rich(s, [
    { text: '+18.66 Followers  Instagram', options: { fontSize:12, bold:true, color:WHITE, breakLine:true } },
    { text: 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam', options: { fontSize:11, color:WHITE } }
  ], { x:8.783, y:2.514, w:2.898, h:0.928, align:'justify', lineSpacingMultiple:1.5, valign:'top' });
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam vulputate dignissim. Enimagi cinec dui', { x:9.044, y:4.424, w:2.898, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam', { x:9.044, y:5.513, w:2.898, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:WHITE });
  rule(s, 9.144, 4.16, 0.681, WHITE, 3);
}

function slide24(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "21_Custom Layout")
  gradRect(s, 6.667, 3.75, 5, 3.75);
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  crosses(s, 7.288, 1.338, 0.675, 'col');
  heading(s, 'Break Slide\n', 'Presentation', { x:1.698, y:2.36, w:4.227, h:1.313, fontSize:36 });
  eyebrow(s, 1.71, 2.149);
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam vulputate dignissim. Enimagi cinec dui nunc', { x:1.71, y:3.808, w:4.629, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'We Take A Rest For 40 Minuets', { x:1.71, y:4.935, w:4.629, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:SUBTLE });
  rule(s, 1.8, 4.728, 4.867, PINK, 2.25);
}

function slide25(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "22_Custom Layout")
  gradRect(s, 0.688, 0.771, 5.021, 6.729);
  imageBox(s, 1.212, 1.696, 5.823, 4.825);
  tagline(s);
  dots(s, 12.665, 5.76);
  txt(s, 'Our Progress', { x:7.289, y:4.074, w:1.622, h:0.37, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:12, bold:true, color:SUBTLE });
  heading(s, 'About Microfest ', 'Computer Device', { x:7.563, y:1.624, w:4.227, h:1.178 });
  eyebrow(s, 7.575, 1.413);
  s.addShape('blockArc', { x:7.663, y:3.163, w:0.874, h:0.874, fill:{color:'F0DAFA'}, line:{type:'none'}, angleRange:[183,128], arcThicknessRatio:0.176 });
  s.addShape('blockArc', { x:7.663, y:3.163, w:0.874, h:0.874, fill:{color:PURPLE}, line:{type:'none'}, angleRange:[56,278], arcThicknessRatio:0.176 });
  txt(s, '71%', { x:7.635, y:3.372, w:0.931, h:0.415, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:SUBTLE });
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam vulputate dignissim. Enimagi cinec', { x:8.717, y:3.36, w:3.022, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  s.addShape('blockArc', { x:7.663, y:4.833, w:0.874, h:0.874, fill:{color:'F0DAFA'}, line:{type:'none'}, angleRange:[183,128], arcThicknessRatio:0.176 });
  s.addShape('blockArc', { x:7.663, y:4.833, w:0.874, h:0.874, fill:{color:PURPLE}, line:{type:'none'}, angleRange:[325,278], arcThicknessRatio:0.176 });
  txt(s, '94%', { x:7.635, y:5.043, w:0.931, h:0.415, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:SUBTLE });
  txt(s, 'Our Progress', { x:7.289, y:5.716, w:1.622, h:0.37, align:'center', lineSpacingMultiple:1.5, valign:'top', fontSize:12, bold:true, color:SUBTLE });
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretiu mquam vulputate dignissim. Enimagi cinec', { x:8.717, y:5.031, w:3.022, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
}

function slide26(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "23_Custom Layout")
  gradRect(s, 6.667, 3.75, 3.042, 3.75);
  imageBox(s, 7.041, 0.829, 5.478, 5.478);
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  heading(s, 'About Microfest ', 'Tablet Device', { x:1.789, y:1.771, w:4.227, h:1.178 });
  eyebrow(s, 1.801, 1.561);
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretium quam vulputate dignissim. Enimagi cinec dui nunc mattis enim ut tell us ma elem entum iena nec dui cinec dui nunc', { x:1.789, y:3.664, w:4.504, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Sharing Idol Experiences', { x:2.062, y:3.166, w:2.463, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:PURPLE });
  s.addShape('triangle', { x:1.886, y:3.33, w:0.153, h:0.132, fill:{color:PURPLE}, line:{type:'none'}, rotate:90 });
  txt(s, 'Aliquaet nibh praesent tristique magna sit amet. Risus pretium quam vulputate dignissim. Enimagi ci', { x:1.801, y:5.313, w:4.504, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Tips And Trick ', { x:2.074, y:4.815, w:2.463, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, bold:true, color:PURPLE });
  s.addShape('triangle', { x:1.898, y:4.98, w:0.153, h:0.132, fill:{color:PURPLE}, line:{type:'none'}, rotate:90 });
}

function slide27(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "24_Custom Layout")
  gradRect(s, 0.688, 0, 5.979, 7.5);
  imageBox(s, 0.521, 0.547, 6.458, 6.458);
  tagline(s);
  dots(s, 12.665, 5.76);
  heading(s, 'About Microfest ', 'Phone Device', { x:7.563, y:1.659, w:4.227, h:1.178 });
  eyebrow(s, 7.575, 1.449);
  txt(s, '2018', { x:7.563, y:3.079, w:1.071, h:0.64, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:24, bold:true, color:GREY });
  txt(s, 'Description Here', { x:8.417, y:3.387, w:2.638, h:0.286, valign:'top', fontSize:11, italic:true, color:GREY });
  txt(s, 'Aliquaet nibh praesent tristique magna sitamet. Risus peretium quavgfm vulputate et nibh praesent magna sitamet. Ris us preti um qu am vulputate et nibh praesent', { x:7.563, y:3.674, w:4.501, h:0.903, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Aliquaet nibh praesent tristique magna sitamet. Risus peretium quavgfm vulputate et nibh praesent magna sitamet. ', { x:7.563, y:4.72, w:4.501, h:0.625, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
  txt(s, 'Next Pages', { x:7.563, y:5.636, w:4.501, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, color:PURPLE });
  icon(s, 'chevron', 8.751, 5.753, 0.141, 0.25, PURPLE);
}

function slide28(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "25_Custom Layout")
  gradRect(s, 6.688, 0, 5.979, 3.75);
  cornerDeco(s, 0, 5.153);
  logo(s, 0.563, 0.577);
  crosses(s, 7.278, 4.3, 0.694, 'col');
  txt(s, 'There’s nothing better than life music. It’s raw energy, and raw energy feeds the soul.', { x:1.422, y:2.762, w:4.95, h:1.851, lineSpacingMultiple:1.5, valign:'top', fontSize:24, italic:true, color:SUBTLE });
  txt(s, '- Dhani Jones', { x:1.432, y:4.873, w:2.557, h:0.37, valign:'top', fontSize:16, bold:true, color:PURPLE });
  txt(s, '“', { x:1.422, y:2.257, w:0.703, h:1.01, valign:'top', fontSize:54, bold:true, color:PURPLE });
}

function slide29(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "26_Custom Layout")
  gradRect(s, 1.646, 3.75, 5.021, 3.75);
  crosses(s, 5.278, 1.231, 0.694, 'col');
  tagline(s);
  dots(s, 12.665, 5.76);
  heading(s, 'About Microfest ', 'Contact Information', { x:7.563, y:1.634, w:4.635, h:1.178 });
  eyebrow(s, 7.575, 1.423);
  icon(s, 'phone', 7.589, 3.286, 0.415, 0.411, PURPLE);
  icon(s, 'envelope', 7.575, 4.368, 0.452, 0.337, PINK);
  txt(s, 'Website', { x:8.225, y:5.361, w:2.194, h:0.404, valign:'top', fontSize:18, bold:true, color:PURPLE });
  txt(s, 'www.microfestlivemusicevent.co', { x:8.225, y:5.724, w:3.343, h:0.352, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY50 });
  txt(s, 'Phones', { x:8.262, y:3.166, w:2.194, h:0.404, valign:'top', fontSize:18, bold:true, color:PURPLE });
  txt(s, '+21 0011 2931 1021/ +21 02 981 102', { x:8.22, y:3.44, w:3.874, h:0.415, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, color:GREY50 });
  txt(s, 'E-Mail', { x:8.238, y:4.265, w:2.194, h:0.404, valign:'top', fontSize:18, bold:true, color:PINK });
  txt(s, 'microfestlivemusic,com', { x:8.238, y:4.535, w:3.874, h:0.421, align:'justify', lineSpacingMultiple:1.5, valign:'top', fontSize:14, color:GREY50 });
  icon(s, 'mobile', 7.654, 5.467, 0.298, 0.513, PURPLE);
}

function slide30(p) {
  const s = p.addSlide();
  // backdrop (from slide layout "31_Title Slide")
  flat(s, 1.625, 0, 3.062, 7.5, PALE);
  gradRect(s, 3.625, 3.75, 3.042, 3.75);
  crosses(s, 5.167, 1.247, 1.042, 'col');
  rich(s, [
    { text: 'Tha', options: { fontSize:96, color:PINK, fontFace:'Raleway Black' } },
    { text: 'nks', options: { fontSize:96, fontFace:'Raleway Black' } }
  ], { x:7.52, y:2.146, w:5.25, h:1.717, valign:'top' });
  txt(s, 'AND SEE YOU NEXT TIME', { x:7.544, y:3.708, w:3.457, h:0.438, valign:'top', wrap:false, fontSize:20, bold:true, color:PINK });
  tagline(s);
  dots(s, 12.665, 5.76);
  arrowCircle(s, 7.62, 4.833, 0.37, GRAD_MID);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit va mus vel euismod leo. Donec commodo et urna', { x:8.094, y:4.705, w:4.231, h:0.625, lineSpacingMultiple:1.5, valign:'top', fontSize:11, color:GREY });
}


// ---------------------------------------------------------------- render
slide01(pres);
slide02(pres);
slide03(pres);
slide04(pres);
slide05(pres);
slide06(pres);
slide07(pres);
slide08(pres);
slide09(pres);
slide10(pres);
slide11(pres);
slide12(pres);
slide13(pres);
slide14(pres);
slide15(pres);
slide16(pres);
slide17(pres);
slide18(pres);
slide19(pres);
slide20(pres);
slide21(pres);
slide22(pres);
slide23(pres);
slide24(pres);
slide25(pres);
slide26(pres);
slide27(pres);
slide28(pres);
slide29(pres);
slide30(pres);

pres.writeFile({ fileName: path.join(__dirname, '10696a22-5567-4aee-b547-ad3d3c6161f1_grok_final.pptx') })
  .then(f => console.log('wrote', f));
