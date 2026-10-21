/**
 * Hoschen - Kitchen Equipment Presentation
 * Standalone pptxgenjs rebuild of the 30-slide reference deck (20 x 11.25 in).
 *
 * Photographs in the original are replaced by flat placeholder rectangles and
 * line-art icons are redrawn from primitives, so the script has no binary assets.
 *
 * Run: node <thisfile>.js   ->  writes <thisfile>.pptx next to the script.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ------------------------------------------------------------------ palette
const OLIVE = '5D5841';      // brand olive
const CREAM = 'FFE7C9';      // outline accent on cards
const GREY = '7F7F7F';       // body copy
const SLATE = '595959';
const MIST = 'F2F2F2';
const INK = '000000';
const WHITE = 'FFFFFF';
const NAVY = '1C3052';
const SMOKE = '494949';
const PHOTO = 'C3C3C3';      // image placeholder tone
const PHOTO_DARK = '5B5B5B';

// -------------------------------------------------------------------- fonts
const POP = 'Poppins';
const POPM = 'Poppins Medium';
const POPS = 'Poppins SemiBold';
const ASI = 'Assistant';
const ASIS = 'Assistant SemiBold';
const ASIX = 'Assistant ExtraBold';

// ------------------------------------------------------------- text presets
const BODY = { fontFace: POP, fontSize: 18, color: GREY, lineSpacingMultiple: 1.5 };
const BODY16 = { fontFace: POP, fontSize: 16, color: GREY, lineSpacingMultiple: 1.5 };
const H1 = { fontFace: POPM, fontSize: 48, color: INK };
const H2 = { fontFace: ASIS, fontSize: 24, color: INK };
const H3 = { fontFace: ASIS, fontSize: 20, color: INK };
const KICKER = { fontFace: ASIS, fontSize: 18, color: OLIVE };
const NAVLINK = { fontFace: ASIS, fontSize: 20, color: GREY, align: 'center' };
const BTN = { fontFace: POPM, fontSize: 20, color: WHITE, align: 'center' };
const STAT = { fontFace: ASI, fontSize: 36, color: INK, bold: true, italic: true };
const PRICE = { fontFace: ASI, fontSize: 36, color: INK, bold: true };
const PLAN = { fontFace: POP, fontSize: 18, color: SLATE, lineSpacingMultiple: 1.5, align: 'center' };
const PRICETAG = { fontFace: POPS, fontSize: 36, color: INK, lineSpacingMultiple: 1.5, align: 'center' };

// Every text frame in the source deck uses the same Google-Slides insets
// (0.1in sides, 0.05in top/bottom). pptxgenjs takes margins in points.
const INSET = [7.2, 7.2, 3.6, 3.6]; // left, right, bottom, top

// ------------------------------------------------------------------ helpers
function shade(blur, offset, color, opacity) {
  return { type: 'outer', angle: 45, blur: blur, offset: offset, color: color, opacity: opacity };
}

/** Filled / outlined rectangle. */
function box(s, [x, y, w, h], opts = {}) {
  s.addShape('rect', { x, y, w, h, ...opts });
}

/** Straight rule; zero width or height gives a pure horizontal/vertical line. */
function rule(s, x, y, w, h, color, width) {
  s.addShape('line', { x, y, w, h, line: { color, width } });
}

/** Text frame. `body` is a plain string or an array of pptxgenjs text runs. */
function txt(s, body, [x, y, w, h], opts = {}) {
  s.addText(body, { x, y, w, h, valign: 'top', margin: INSET, wrap: true, ...opts });
}

/** Placeholder standing in for a photograph in the reference deck. */
function photo(s, [x, y, w, h], opts = {}) {
  s.addShape('rect', { x, y, w, h, fill: PHOTO, ...opts });
}

/**
 * Chrome of the phone / laptop / desktop mockups; proportions follow the
 * silhouettes in the reference art. photo() paints the screen afterwards.
 */
function device(s, [x, y, w, h], kind) {
  const CASE = 'CFCFCF';
  const SHELL = '1F1F1F';
  if (kind === 'phone') {
    s.addShape('roundRect', { x, y, w, h, rectRadius: h * 0.14, fill: SHELL });
  } else if (kind === 'laptop') {
    s.addShape('rect', { x: x + w * 0.093, y, w: w * 0.815, h: h * 0.93, fill: SHELL });
    s.addShape('roundRect', { x, y: y + h * 0.93, w, h: h * 0.07, rectRadius: 0.07, fill: CASE });
  } else {
    s.addShape('roundRect', { x: x + w * 0.018, y: y + h * 0.03, w: w * 0.964, h: h * 0.76,
                              rectRadius: 0.09, fill: SHELL });
    s.addShape('trapezoid', { x: x + w * 0.4, y: y + h * 0.79, w: w * 0.2, h: h * 0.15,
                              fill: CASE, flipV: true });
    s.addShape('roundRect', { x: x + w * 0.33, y: y + h * 0.93, w: w * 0.34, h: h * 0.055,
                              rectRadius: 0.05, fill: CASE });
  }
}

/**
 * Storefront navigation header shared by the cover, section and closing slides.
 * `active` is the index of the link rendered in black (0 = none highlighted).
 */
const NAV_LINKS = ['Latest Products', 'Our Collection', 'About Us', 'Customer Services '];

function navbar(s, active = 0) {
  txt(s, 'HOSCHEN', [1.166, 0.498, 2.286, 0.505], { ...H1, fontSize: 24, color: OLIVE, bold: true });
  NAV_LINKS.forEach((label, i) => {
    txt(s, label, [4.287 + i * 2.7855, 0.531, 2.592, 0.438],
        { ...NAVLINK, color: active === i + 1 ? INK : GREY });
  });
  glyph(s, [17.65, 0.531, 0.438, 0.438], 'basket', OLIVE);       // cart
  s.addShape('ellipse', { x: 16.904, y: 0.6, w: 0.262, h: 0.262, line: { color: OLIVE, width: 0.75 } });
  rule(s, 17.12, 0.823, 0.064, 0.077, OLIVE, 0.75);              // search handle
  glyph(s, [18.534, 0.6, 0.3, 0.3], 'person', OLIVE);            // account
}

/**
 * Five-bar "brush stroke" motif. Bars stack along `axis` over `span` inches and
 * alternate between two columns `dx` apart; `thick` is the cross-axis size.
 */
const BRUSH_BARS = [[0, 0.153, 0], [0.2358, 0.153, 1], [0.4335, 0.1585, 0],
                    [0.6439, 0.147, 1], [0.8474, 0.153, 0]];

function brush(s, x, y, thick, dx, span, color, opts = {}) {
  const { axis = 'y', rows = [0, 1, 2, 3, 4], flip = false, swap = false, transparency } = opts;
  const end = BRUSH_BARS[4][0] + BRUSH_BARS[4][1];
  let bars = rows.map(i => BRUSH_BARS[i]);
  if (flip) bars = rows.map(i => BRUSH_BARS[4 - i]).map(([v, ln, c]) => [end - v - ln, ln, c]).reverse();
  const v0 = bars[0][0];
  const fill = transparency ? { color, transparency } : color;
  bars.forEach(([v, ln, c]) => {
    const off = (v - v0) * span;
    const len = ln * span;
    const cross = (swap ? 1 - c : c) * dx;
    if (axis === 'y') s.addShape('rect', { x: x + cross, y: y + off, w: thick, h: len, fill });
    else s.addShape('rect', { x: x + off, y: y + cross, w: len, h: thick, fill });
  });
}

/** Line-art stand-ins for the reference deck's recurring icon images. */
function glyph(s, [x, y, w, h], kind, color) {
  const ln = { color, width: 1.25 };
  if (kind === 'plus') {
    s.addShape('rect', { x: x + w * 0.06, y: y + h * 0.44, w: w * 0.88, h: h * 0.12, fill: color });
    s.addShape('rect', { x: x + w * 0.44, y: y + h * 0.06, w: w * 0.12, h: h * 0.88, fill: color });
  } else if (kind === 'uturn') {
    // The source art sits inside a transparent margin; match its visible extent
    // and quarter-turn it clockwise so the head points left.
    const aw = w * 0.79, ah = h * 0.51;
    const cx = x + w * 0.5, cy = y + h * 0.5;
    s.addShape('uturnArrow', { x: cx - ah / 2, y: cy - aw / 2, w: ah, h: aw,
                               rotate: 90, fill: WHITE, line: { color, width: 2.5 } });
  } else if (kind === 'basket') {
    s.addShape('trapezoid', { x: x + w * 0.08, y: y + h * 0.42, w: w * 0.84, h: h * 0.4, flipV: true, line: ln });
    s.addShape('arc', { x: x + w * 0.28, y: y + h * 0.22, w: w * 0.44, h: h * 0.4,
                        angleRange: [180, 0], line: ln });
  } else if (kind === 'cart') {
    s.addShape('rect', { x: x + w * 0.2, y: y + h * 0.22, w: w * 0.62, h: h * 0.42, line: ln });
    s.addShape('ellipse', { x: x + w * 0.28, y: y + h * 0.72, w: w * 0.14, h: h * 0.14, fill: color });
    s.addShape('ellipse', { x: x + w * 0.6, y: y + h * 0.72, w: w * 0.14, h: h * 0.14, fill: color });
  } else if (kind === 'person') {
    s.addShape('ellipse', { x, y, w, h, line: ln });
    s.addShape('ellipse', { x: x + w * 0.34, y: y + h * 0.22, w: w * 0.32, h: h * 0.32, line: ln });
    s.addShape('arc', { x: x + w * 0.22, y: y + h * 0.58, w: w * 0.56, h: h * 0.5,
                        angleRange: [180, 0], line: ln });
  }
}

/**
 * Pictogram placeholder: a rounded tile of the icon's own colour, which reads as
 * the small mono icons (cart, chef hat, tap, ...) used throughout the deck.
 */
function icon(s, [x, y, w, h], color) {
  const p = Math.min(w, h) * 0.14;
  s.addShape('roundRect', {
    x: x + p, y: y + p, w: w - 2 * p, h: h - 2 * p,
    rectRadius: Math.min(w, h) * 0.14, line: { color, width: 1.25 },
  });
}


// ------------------------------------------------- repeated placeholder copy
const LOREM1 =
  'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor adipiscing turpis.';
const LOREM2 =
  'Lorem sed risus ultricies tristique nulla aliquet enim.';
const LOREM3 =
  'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor.';
const LOREM4 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. aliqua. Nibh tortor id nisl purus in mollis nunc sed id. Posuere morbi urna molestie at eu facilisis sed. Condimentum mattis lectus.';
const LOREM5 =
  'Lorem ipsum dolor sit amet, consectetur elit. Nibh morbi urna molestie placerat  ullamcorper magna placerat.';
const LOREM6 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. this aliqua. Nibh posuere morbi leo urna molestie at nibh in ipsum nisl ullamcorper magna ac placerat vestibulum lectus mauris ';
const LOREM7 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. this aliqua. Eget mi proin sed libero enim sed. met porttitor dolor.';
const LOREM8 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod.';
const LOREM9 =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit,.';
const LOREM10 =
  'Lorem sed risus ultricies tristique nulla aliquet';

// ------------------------------------------------------------------ slides
function slide01(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 5.857, 20, 5.476], { fill: OLIVE });
  rule(s, 0, 1.5, 20, 0, OLIVE, 2.25);
  txt(s, 'Since 2010', [1.17, 2.371, 2.477, 0.505], { ...H2 });
  txt(s, 'Hoschen', [1.17, 3.034, 8.83, 1.952], { ...H1, fontSize: 110, color: OLIVE });
  txt(s, 'Kitchen Equipment Presentation', [11.489, 2.781, 7.044, 0.505], { ...H2, align: 'right' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et.', [11.489, 3.609, 7.044, 0.967], { ...BODY, align: 'right' });
  navbar(s, 0);
  photo(s, [0.518, 6.424, 18.964, 4.342]);
}

function slide02(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 0, 7.786, 11.25], { fill: OLIVE });
  box(s, [0.538, 0.503, 6.711, 10.245], { line: { color: WHITE, width: 2.25 } });
  txt(s, 'Welcome to Hoschen', [10, 2.28, 8.51, 0.909], { ...H1 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. aliqua. Nibh tortor id aliquet lectus. Amet nisl purus in mollis nunc sed id. Posuere morbi leo urna molestie at elementum eu facilisis sed. Condimentum mattis sed pellentesque id nibh tortor id aliquet lectus. Vivamus arcu felis ets bibendum ut tristique et. Mattis ullamcorper velit sed pretium. Blandit libero volutpat sed cras. Nibh cras pulvinar mattis nunc sed blandit. Mauris commodo quis imperdiet massa lacus non odio.', [10, 4.283, 8.8, 3.693], { ...BODY });
  txt(s, 'Delivery From 3 Weeks', [10.733, 3.515, 5.506, 0.37], { ...BODY16, lineSpacingMultiple: 1 });
  icon(s, [10, 3.445, 0.512, 0.512], OLIVE);
  txt(s, 'ABOUT US', [10, 1.677, 3.414, 0.404], { ...KICKER });
  box(s, [10, 8.569, 3.261, 1.004], { fill: OLIVE, line: { color: NAVY, width: 1 } });
  txt(s, 'Add to Cart', [10.915, 8.852, 1.997, 0.438], { ...BTN });
  glyph(s, [10.349, 8.875, 0.392, 0.392], 'cart', WHITE);
  box(s, [13.529, 8.569, 2.841, 1.004], { line: { color: OLIVE, width: 1.5 } });
  txt(s, 'Learn More', [13.951, 8.852, 1.997, 0.438], { ...BTN, color: INK });
  box(s, [16.639, 8.569, 1.078, 1.004], { line: { color: OLIVE, width: 1.5 } });
  glyph(s, [16.97, 8.863, 0.415, 0.415], 'plus', OLIVE);
  photo(s, [1.2, 1.082, 5.386, 9.086]);
}

function slide03(pptx) {
  const s = pptx.addSlide();
  box(s, [11.91, 1.24, 6.78, 8.771], { fill: OLIVE });
  photo(s, [12.648, 1.955, 6.042, 7.339]);
  box(s, [12.648, 1.955, 0.481, 7.339], { fill: WHITE });
  box(s, [12.648, 8.833, 6.042, 0.461], { fill: WHITE });
  brush(s, 12.821, 3.042, 0.044, 0.066, 0.865, OLIVE, { swap: true });
  brush(s, 12.821, 5.193, 0.044, 0.066, 0.864, OLIVE, { swap: true });
  brush(s, 12.821, 7.343, 0.044, 0.066, 0.864, OLIVE, { swap: true });
  brush(s, 13.24, 9.02, 0.044, 0.067, 0.864, OLIVE, { axis: 'x' });
  brush(s, 15.39, 9.02, 0.044, 0.067, 0.864, OLIVE, { axis: 'x' });
  brush(s, 17.54, 9.02, 0.044, 0.067, 0.865, OLIVE, { axis: 'x' });
  txt(s, 'Vision And Mission', [1.31, 2.225, 7.6, 0.909], { ...H1 });
  txt(s, 'ABOUT US', [1.31, 1.622, 3.414, 0.404], { ...KICKER });
  txt(s, 'Vision', [1.31, 3.799, 5.643, 0.505], { ...H2 });
  txt(s, LOREM4, [1.31, 4.505, 7.793, 1.876], { ...BODY });
  txt(s, 'Mission', [1.31, 7.046, 5.643, 0.505], { ...H2 });
  txt(s, LOREM4, [1.31, 7.752, 7.793, 1.876], { ...BODY });
}

function slide04(pptx) {
  const s = pptx.addSlide();
  box(s, [5.761, 3.599, 2.748, 0.127], { fill: { color: SMOKE, transparency: 50 } });
  box(s, [5.761, 3.246, 2.748, 0.132], { fill: { color: SMOKE, transparency: 50 } });
  box(s, [1.64, 3.775, 2.748, 0.132], { fill: OLIVE });
  box(s, [1.64, 3.417, 2.748, 0.137], { fill: OLIVE });
  box(s, [1.64, 3.042, 2.748, 0.132], { fill: OLIVE });
  box(s, [1.64, 5.749, 2.748, 0.127], { fill: { color: SMOKE, transparency: 50 }, flipH: true });
  box(s, [1.64, 5.396, 2.748, 0.132], { fill: { color: SMOKE, transparency: 50 }, flipH: true });
  box(s, [5.761, 5.925, 2.748, 0.132], { fill: { color: SMOKE, transparency: 50 }, flipH: true });
  box(s, [5.761, 5.567, 2.748, 0.137], { fill: { color: SMOKE, transparency: 50 }, flipH: true });
  box(s, [5.761, 5.193, 2.748, 0.132], { fill: { color: SMOKE, transparency: 50 }, flipH: true });
  box(s, [5.761, 7.899, 2.748, 0.127], { fill: OLIVE });
  box(s, [5.761, 7.546, 2.748, 0.132], { fill: OLIVE });
  box(s, [1.64, 8.075, 2.748, 0.132], { fill: OLIVE });
  box(s, [1.64, 7.718, 2.748, 0.137], { fill: OLIVE });
  box(s, [1.64, 7.343, 2.748, 0.132], { fill: OLIVE });
  box(s, [0, 10.288, 20, 0.962], { fill: OLIVE });
  box(s, [0, 0.002, 20, 0.962], { fill: OLIVE });
  txt(s, 'Kitchen Of Your Inspiration', [11.075, 2.763, 6.977, 1.717], { ...H1 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore aliqua. Morbi leo urna molestie at eu facilisis ets sed consequat Condimentum mattis sed lectus elementum donec.', [11.075, 4.855, 7.285, 1.876], { ...BODY });
  txt(s, 'ABOUT US', [11.075, 2.16, 3.414, 0.404], { ...KICKER });
  txt(s, 'Happy Customers', [13.432, 7.387, 4.621, 0.505], { ...H2 });
  txt(s, 'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor', [13.432, 8.123, 4.621, 0.967], { ...BODY });
  txt(s, '+750K', [11.075, 7.911, 1.887, 0.707], { ...STAT });
  photo(s, [2.417, 2.701, 5.314, 5.847]);
}

function slide05(pptx) {
  const s = pptx.addSlide();
  box(s, [11.092, 5.251, 2.748, 0.132], { fill: OLIVE });
  box(s, [11.092, 4.893, 2.748, 0.137], { fill: OLIVE });
  box(s, [11.092, 4.519, 2.748, 0.132], { fill: OLIVE });
  box(s, [11.619, 0, 8.381, 6.357], { fill: OLIVE });
  txt(s, 'We Supply & Serve Kitchen Solution', [1.378, 2.365, 7.5, 1.717], { ...H1 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. aliqua. Nibh tortor id aliquet lectus. Amet nisl purus in mollis nunc sed id. Posuere morbi leo urna molestie at elementum eu pulvinar mattis nunc sed blandit. Mauris commodo quis imperdiet massa lacus non odio. Amet commodo nulla facilisi nullam vehicula ipsum a. Sem integer vitae justo consequat nisl.', [1.384, 4.59, 7.707, 3.239], { ...BODY });
  txt(s, 'ABOUT US', [1.378, 1.763, 3.414, 0.404], { ...KICKER });
  box(s, [1.378, 8.483, 2.943, 1.004], { fill: OLIVE });
  txt(s, 'Learn More', [1.851, 8.766, 1.997, 0.438], { ...BTN });
  box(s, [4.589, 8.483, 1.078, 1.004], { line: { color: OLIVE, width: 1.5 } });
  glyph(s, [4.921, 8.778, 0.415, 0.415], 'plus', OLIVE);
  txt(s, 'Famous Brand', [14.154, 7.773, 4.035, 0.505], { ...H2 });
  txt(s, 'Lorem ipsum dolor sit amet, elit, sed do ets eiusmod.', [14.154, 8.521, 4.035, 0.967], { ...BODY });
  txt(s, '150+', [12.052, 8.277, 1.732, 0.707], { ...STAT });
  box(s, [11.619, 7.248, 7.003, 2.765], { line: { color: CREAM, width: 3 } });
  photo(s, [14.762, 3.545, 5.238, 2.813]);
  photo(s, [11.94, 0, 7.738, 3.262]);
  photo(s, [11.619, 3.545, 2.81, 2.813]);
}

function slide06(pptx) {
  const s = pptx.addSlide();
  box(s, [15.801, 10.053, 2.748, 0.132], { fill: OLIVE });
  box(s, [15.801, 6.328, 2.748, 0.132], { fill: OLIVE });
  box(s, [1.452, 10.053, 2.748, 0.132], { fill: OLIVE });
  box(s, [1.452, 6.328, 2.748, 0.132], { fill: OLIVE });
  box(s, [1.85, 7.004, 2.317, 2.511], { fill: { color: OLIVE, transparency: 60 } });
  box(s, [1.452, 6.587, 7.905, 3.345], { line: { color: OLIVE, width: 1 } });
  txt(s, 'Life Is Beautiful Through Cooking', [3.876, 1.667, 12.248, 0.909], { ...H1, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. aliqua. Nibh tortor id aliquet lectus. Amet nisl purus in mollis nunc sed id. Posuere morbi leo urna molestie at elementum eu pulvinar mattis nunc sed blandit. Mauris commodo quis imperdiet massa lacus non odio. Nulla facilisi nullam vehicula ipsum a. Sem integer vitae justo consequat nisl. Et malesuada fames ac turpis. Amet commodo nulla facilisi nullam vehicula ipsum a. Sem integer vitae justo eget magna fermentum.', [1.452, 2.993, 17.097, 1.876], { ...BODY, align: 'center' });
  txt(s, 'ABOUT US', [8.293, 1.065, 3.414, 0.404], { ...KICKER, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor', [5.002, 7.866, 3.76, 1.275], { ...BODY16 });
  txt(s, 'Quality Products', [5.002, 7.231, 3.76, 0.438], { ...H3 });
  photo(s, [1.85, 7.004, 2.317, 2.511]);
  photo(s, [10.643, 6.587, 7.905, 3.345]);
  txt(s, '98%', [2.306, 7.906, 1.405, 0.707], { ...STAT, color: WHITE, align: 'center' });
}

function slide07(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 0, 20, 11.25], { line: { color: OLIVE, width: 20 } });
  brush(s, 19.273, 0.568, 0.727, 0, 1.466, OLIVE, { rows: [1, 3] });
  brush(s, 0, 0.568, 0.727, 0, 1.466, OLIVE, { rows: [1, 3] });
  brush(s, 19.273, 9.868, 0.727, 0, 1.468, OLIVE, { rows: [1, 3] });
  brush(s, 0, 9.868, 0.727, 0, 1.468, OLIVE, { rows: [1, 3] });
  txt(s, 'Estimated Earning', [7.089, 7.79, 5.822, 0.438], { ...NAVLINK, fontFace: POPS, color: OLIVE });
  txt(s, '$478,34 ', [8.791, 9.176, 2.417, 0.64], { ...H1, fontFace: ASIX, fontSize: 32, align: 'center' });
  txt(s, 'IN JUNE 2024', [8.791, 8.575, 2.417, 0.37], { ...PRICETAG, fontSize: 16, lineSpacingMultiple: 1 });
  txt(s, 'PRODUCT SOLD', [3.932, 8.575, 2.417, 0.37], { ...PRICETAG, fontSize: 16, lineSpacingMultiple: 1 });
  txt(s, '+400K ', [4.048, 9.176, 2.063, 0.64], { ...H1, fontFace: ASIX, fontSize: 32, align: 'center' });
  txt(s, '$1851,86 ', [13.65, 9.176, 2.417, 0.64], { ...H1, fontFace: ASIX, fontSize: 32, align: 'center' });
  txt(s, 'THIS MONTH', [13.773, 8.575, 2.173, 0.37], { ...PRICETAG, fontSize: 16, lineSpacingMultiple: 1 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. this aliqua. Nibh morbi leo urna molestie at nibh in ipsum nisl ullamcorper magna ac placerat vestibulum lectus mauris elementum eu. feugiat pretium. lorem Phasellus vestibulum lorem sed risus ultricies tristique nulla aliquet enim. Ullamcorper a lacus vestibulum sed arcu non odio euismod. Massa tincidunt nunc pulvinar sapien et ligula. Ipsum dolor sit amet consectetur adipiscing elit duis. Nibh cras pulvinar mattis nunc sed blandit Mattis ullamcorper velit sed ullamcorper malesuada fames ac turpis posuere. sed risus ultricies tristique.', [1.607, 4.367, 16.785, 2.33], { ...BODY, align: 'center' });
  txt(s, 'We Work With Leading Brands A Cross The Country', [4.682, 2.037, 10.636, 1.717], { ...H1, align: 'center' });
  txt(s, 'ABOUT US', [8.293, 1.435, 3.414, 0.404], { ...KICKER, align: 'center' });
}

function slide08(pptx) {
  const s = pptx.addSlide();
  box(s, [-0.014, 0, 20.014, 3.821], { fill: OLIVE });
  box(s, [5.577, 1.342, 8.846, 4.283], { fill: WHITE, line: { color: WHITE, width: 3 } });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing.', [14.792, 8.33, 3.654, 0.871], { ...BODY16 });
  txt(s, 'Modern Kitchen Equipment ', [14.792, 7.695, 3.654, 0.438], { ...H3 });
  txt(s, 'Kitchen Inspiration', [1.554, 7.127, 7.5, 0.909], { ...H1 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. aliqua. Nibh posuere morbi leo urna molestie at elementum eu. feugiat pretium nibh ipsum consequat nisl ullamcorper.', [1.554, 8.495, 7.707, 1.876], { ...BODY });
  txt(s, 'ABOUT US', [1.554, 6.525, 3.414, 0.404], { ...KICKER });
  photo(s, [6.231, 1.996, 7.538, 3.65]);
  photo(s, [11.53, 7.493, 2.317, 1.909]);
}

function slide09(pptx) {
  const s = pptx.addSlide();
  box(s, [4.869, 1.276, 4.816, 3.368], { fill: OLIVE });
  box(s, [13.733, 1.276, 4.816, 3.368], { fill: OLIVE });
  txt(s, 'The Heart of The Kitchen', [5.606, 6.299, 8.788, 0.909], { ...H1, align: 'center' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. aliqua. Nibh tortor id aliquet lectus. Amet nisl purus in mollis nunc sed id. Posuere morbi leo urna molestie at elementum eu pulvinar mattis nunc sed blandit. Mauris commodo quis imperdiet massa lacus non odio. Nulla facilisi nullam vehicula ipsum a. Sem integer vitae justo consequat nisl. Et', [1.452, 7.625, 17.097, 1.421], { ...BODY, align: 'center' });
  txt(s, 'ABOUT US', [8.293, 5.696, 3.414, 0.404], { ...KICKER, align: 'center' });
  txt(s, 'Modern Design', [6.342, 9.57, 3.586, 0.404], { ...KICKER, color: INK, align: 'center' });
  txt(s, 'Top In The Market', [10.073, 9.57, 3.586, 0.404], { ...KICKER, color: INK, align: 'center' });
  txt(s, 'Quality Products', [13.803, 9.57, 3.586, 0.404], { ...KICKER, color: INK, align: 'center' });
  txt(s, 'Elegant & Engineered', [2.611, 9.57, 3.586, 0.404], { ...KICKER, color: INK, align: 'center' });
  box(s, [11.175, 1.672, 1.078, 1.004], { fill: OLIVE, line: { color: OLIVE, width: 1 } });
  txt(s, 'Minimalist Kitchen Equipment', [10.438, 3.19, 2.541, 1.058], { ...H3, lineSpacingMultiple: 1.5, align: 'center' });
  icon(s, [11.464, 1.924, 0.5, 0.5], WHITE);
  box(s, [2.188, 1.672, 1.078, 1.004], { fill: OLIVE, line: { color: OLIVE, width: 1 } });
  txt(s, 'Minimalist Kitchen Equipment', [1.452, 3.19, 2.541, 1.058], { ...H3, lineSpacingMultiple: 1.5, align: 'center' });
  icon(s, [2.477, 1.924, 0.5, 0.5], WHITE);
  photo(s, [5.189, 1.573, 4.176, 2.774], { line: { color: WHITE, width: 3 } });
  photo(s, [14.052, 1.573, 4.176, 2.774], { line: { color: WHITE, width: 3 } });
}

function slide10(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 0, 0.896, 11.25], { fill: OLIVE });
  box(s, [0, 10.392, 20, 0.858], { fill: OLIVE });
  brush(s, 0.486, 0.675, 0.541, 0, 1.484, WHITE, { rows: [0, 2, 4] });
  brush(s, 0.486, 4.368, 0.541, 0, 1.483, WHITE, { rows: [0, 2, 4] });
  brush(s, 0.486, 8.06, 0.541, 0, 1.483, WHITE, { rows: [0, 2, 4] });
  brush(s, 17.682, 10.208, 0.63, 0, 1.484, WHITE, { axis: 'x', rows: [0, 2, 4] });
  box(s, [19.166, 0.733, 1.32, 0.17], { fill: WHITE });
  box(s, [19.166, 0.262, 1.32, 0.176], { fill: WHITE });
  txt(s, 'Renald Nick Moserry', [2.293, 1.664, 7.5, 0.909], { ...H1 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. aliqua. Nibh posuere morbi leo urna molestie at elementum eu. feugiat pretium nibh ipsum nisl ullamcorper magna ac placerat vestibulum lectus mauris.', [2.293, 3.032, 8.922, 1.876], { ...BODY });
  txt(s, 'EXECUTIVE CHEF', [2.293, 1.062, 3.414, 0.404], { ...KICKER });
  txt(s, 'Write Skill Here', [2.293, 5.454, 4.822, 0.505], { ...H2, bold: true });
  txt(s, LOREM1, [2.293, 6.161, 5.398, 0.967], { ...BODY });
  txt(s, '98%', [9.696, 5.938, 1.405, 0.707], { ...STAT });
  txt(s, '95%', [9.696, 8.158, 1.405, 0.707], { ...STAT });
  txt(s, 'Write Skill Here', [2.293, 7.675, 4.822, 0.505], { ...H2, bold: true });
  txt(s, LOREM1, [2.293, 8.382, 5.398, 0.967], { ...BODY });
  photo(s, [13.762, 0, 6.238, 9.09]);
}

function slide11(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 7.69, 9.262, 3.56], { fill: OLIVE, flipH: true });
  photo(s, [0, 0, 4.631, 7.69]);
  photo(s, [4.631, 0, 4.631, 7.69], { fill: GREY });
  txt(s, 'Chef De Partie', [1.35, 9.317, 6.563, 0.909], { ...H1, color: WHITE });
  txt(s, 'OUR TEAM', [1.35, 8.715, 3.414, 0.404], { ...KICKER, color: MIST });
  txt(s, LOREM5, [11.283, 2.628, 7.508, 0.967], { ...BODY });
  txt(s, 'Kennyl Zack Jossie', [12.508, 1.86, 5.271, 0.505], { ...PRICE, fontSize: 24 });
  txt(s, '01.', [11.283, 1.759, 1.051, 0.707], { ...STAT });
  txt(s, LOREM5, [11.283, 4.965, 7.508, 0.967], { ...BODY });
  txt(s, 'Barret Jacono Lihu', [12.508, 4.196, 5.271, 0.505], { ...PRICE, fontSize: 24 });
  txt(s, '02.', [11.283, 4.095, 1.051, 0.707], { ...STAT });
  txt(s, 'HOSCHEN', [14.752, 8.465, 4.038, 0.909], { ...H1, color: OLIVE, bold: true, align: 'right' });
  rule(s, 15.527, 9.964, 3.191, 0, OLIVE, 6);
  txt(s, '02', [7.912, 0.79, 1.051, 0.707], { ...STAT, color: WHITE });
  rule(s, 8.438, -0.061, 0, 0.608, WHITE, 6);
  txt(s, '01', [3.281, 0.79, 1.051, 0.707], { ...STAT, color: WHITE });
  rule(s, 3.807, -0.061, 0, 0.608, WHITE, 6);
  box(s, [12.221, 8.497, 0.697, 0.649], { fill: OLIVE, line: { color: OLIVE, width: 1 } });
  box(s, [13.159, 8.497, 0.697, 0.649], { fill: OLIVE, line: { color: OLIVE, width: 1 } });
  box(s, [11.283, 8.497, 0.697, 0.649], { fill: OLIVE, line: { color: OLIVE, width: 1 } });
  txt(s, 'Lorem ipsum dolor sit amet, elit, sed do', [11.283, 9.548, 2.736, 0.871], { ...BODY16 });
  icon(s, [11.414, 8.605, 0.434, 0.434], WHITE);
  icon(s, [13.303, 8.617, 0.41, 0.41], WHITE);
  icon(s, [12.364, 8.616, 0.411, 0.411], WHITE);
}

function slide12(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 0, 1.143, 1], { fill: OLIVE });
  photo(s, [9.48, 1.848, 3.777, 6.402], { shadow: shade(15, 15, GREY, 0.4), fill: PHOTO_DARK });
  txt(s, 'Demiano Juanes', [1.493, 3.496, 6.494, 0.909], { ...H1 });
  txt(s, 'COMMIS TEAM', [1.493, 4.637, 3.547, 0.404], { ...KICKER });
  txt(s, 'Lorem ipsum dolor sit amet, elit, sed do ni labore et. aliqua. Nibh posuere morbi leo urna molestie at elementum eu. feugiat pretium nibh ipsum ac placerat vestibulum lectus this amet commodo nulla facilisi nullam vehicula ipsum a. Sem vitae justo eget magna fermentum nisl ullamcorper magna tempor incididunt. fames ac turpis. ', [1.493, 5.333, 6.494, 3.239], { ...BODY });
  txt(s, '01', [1.493, 2.679, 1.092, 0.707], { ...STAT, color: GREY });
  box(s, [9.47, 1.848, 0.544, 6.402], { fill: CREAM });
  txt(s, [
    { text: '02', options: { fontFace: POP, fontSize: 20, color: GREY, bold: true, italic: true } },
    { text: '. Arthur Zain,', options: { ...BODY, fontSize: 20, bold: true, lineSpacingMultiple: 1 } },
    { text: ' 03', options: { fontFace: POP, fontSize: 20, color: GREY, bold: true, italic: true } },
    { text: '. Ashton Nolan Jarrouj, ', options: { ...BODY, fontSize: 20, bold: true, lineSpacingMultiple: 1 } },
    { text: '04', options: { fontFace: POP, fontSize: 20, color: GREY, bold: true, italic: true } },
    { text: '. Vladimir Rexlee ', options: { ...BODY, fontSize: 20, bold: true, lineSpacingMultiple: 1 } },
  ], [9.47, 8.972, 9.683, 0.438]);
  txt(s, '01', [12.139, 2.677, 1.051, 0.707], { ...STAT, color: WHITE });
  rule(s, 12.665, 1.826, 0, 0.608, WHITE, 6);
  txt(s, '02', [14.641, 7.149, 0.76, 0.572], { ...STAT, fontSize: 28, color: WHITE });
  rule(s, 15.021, 7.815, 0, 0.439, WHITE, 6);
  txt(s, '03', [16.889, 7.149, 0.76, 0.572], { ...STAT, fontSize: 28, color: WHITE });
  rule(s, 17.269, 7.815, 0, 0.439, WHITE, 6);
  txt(s, '04', [19.136, 7.149, 0.76, 0.572], { ...STAT, fontSize: 28, color: WHITE });
  rule(s, 19.516, 7.815, 0, 0.439, WHITE, 6);
  photo(s, [15.662, 1.844, 2.09, 6.402], { fill: GREY });
  photo(s, [17.91, 1.841, 2.09, 6.402]);
  photo(s, [13.415, 1.848, 2.09, 6.402]);
}

function slide13(pptx) {
  const s = pptx.addSlide();
  box(s, [9.081, 1.216, 9.625, 8.819], { line: { color: OLIVE, width: 1 } });
  box(s, [1.293, 1.216, 5.873, 8.819], { fill: OLIVE });
  txt(s, 'Write Skill Here', [10.117, 3.189, 4.822, 0.505], { ...H2, bold: true });
  txt(s, LOREM1, [10.117, 3.869, 5.398, 0.967], { ...BODY });
  txt(s, '95%', [16.266, 3.659, 1.405, 0.707], { ...STAT });
  txt(s, 'Professional Chef', [10.117, 1.978, 7.554, 0.64], { ...H1, fontSize: 32 });
  txt(s, 'Write Skill Here', [10.117, 5.407, 4.822, 0.505], { ...H2, bold: true });
  txt(s, LOREM1, [10.117, 6.087, 5.398, 0.967], { ...BODY });
  txt(s, '90%', [16.266, 5.877, 1.405, 0.707], { ...STAT });
  txt(s, 'Write Skill Here', [10.117, 7.625, 4.822, 0.505], { ...H2, bold: true });
  txt(s, LOREM1, [10.117, 8.305, 5.398, 0.967], { ...BODY });
  txt(s, '98%', [16.266, 8.095, 1.405, 0.707], { ...STAT });
  photo(s, [1.293, 1.216, 5.873, 8.819]);
  txt(s, 'Hudsond Bram Leonard', [2.076, 8.033, 3.874, 1.313], { ...BTN, fontSize: 36, align: 'left' });
  txt(s, 'EXECUTIVE SOUS CHEF', [2.076, 7.465, 3.547, 0.404], { ...KICKER, color: WHITE });
}

function slide14(pptx) {
  const s = pptx.addSlide();
  box(s, [1.484, 1.073, 6.894, 2.765], { fill: OLIVE });
  txt(s, '01. Write Title Here', [2.447, 1.578, 4.822, 0.505], { ...H2, color: WHITE });
  txt(s, LOREM1, [2.447, 2.366, 5.398, 0.967], { ...BODY, color: MIST });
  box(s, [1.484, 4.242, 6.894, 2.765], { fill: OLIVE });
  txt(s, '02. Write Title Here', [2.447, 4.748, 4.822, 0.505], { ...H2, color: WHITE });
  txt(s, LOREM1, [2.447, 5.535, 5.398, 0.967], { ...BODY, color: MIST });
  box(s, [1.484, 7.412, 6.894, 2.765], { fill: OLIVE });
  txt(s, '03. Write Title Here', [2.447, 7.917, 4.822, 0.505], { ...H2, color: WHITE });
  txt(s, LOREM1, [2.447, 8.705, 5.398, 0.967], { ...BODY, color: MIST });
  brush(s, 1.31, 1.713, 0.541, 0, 1.484, WHITE, { rows: [0, 2, 4] });
  brush(s, 1.31, 4.883, 0.541, 0, 1.483, WHITE, { rows: [0, 2, 4] });
  brush(s, 1.297, 8.052, 0.541, 0, 1.484, WHITE, { rows: [0, 2, 4] });
  txt(s, [
    { text: LOREM6, options: { ...BODY, breakLine: true } },
    { text: '', options: { ...BODY, color: INK, breakLine: true } },
    { text: LOREM2, options: { ...BODY, bullet: { characterCode: '2751', indent: 22.5 }, breakLine: true } },
    { text: LOREM2, options: { ...BODY, bullet: { characterCode: '2751', indent: 22.5 }, breakLine: true } },
    { text: LOREM2, options: { ...BODY, bullet: { characterCode: '2751', indent: 22.5 } } },
  ], [11.078, 3.936, 7.624, 3.693]);
  txt(s, 'Our Amazing Service', [11.078, 2.569, 7.624, 0.909], { ...H1 });
  txt(s, 'OUR SERVICE', [11.078, 1.966, 3.414, 0.404], { ...KICKER });
  box(s, [11.078, 8.28, 2.964, 1.004], { fill: OLIVE });
  txt(s, 'View Product', [11.32, 8.563, 2.481, 0.438], { ...BTN });
  box(s, [14.268, 8.268, 1.078, 1.016], { line: { color: OLIVE, width: 1.5 } });
  glyph(s, [14.6, 8.569, 0.415, 0.415], 'plus', OLIVE);
}

function slide15(pptx) {
  const s = pptx.addSlide();
  s.background = { color: OLIVE };
  box(s, [1.064, 0.824, 17.872, 9.603], { fill: WHITE });
  box(s, [2.218, 3.287, 6.894, 2.765], { fill: OLIVE, shadow: shade(30, 15, SLATE, 0.4) });
  txt(s, '01. Write Title Here', [3.18, 3.792, 4.822, 0.505], { ...H2, color: WHITE });
  txt(s, LOREM1, [3.18, 4.58, 5.398, 0.967], { ...BODY, color: MIST });
  box(s, [11.176, 3.287, 6.894, 2.765], { line: { color: CREAM, width: 3 } });
  txt(s, '02. Write Title Here', [12.138, 3.792, 4.822, 0.505], { ...H2 });
  txt(s, LOREM1, [12.138, 4.58, 5.398, 0.967], { ...BODY });
  box(s, [11.234, 6.581, 6.894, 2.765], { line: { color: CREAM, width: 3 } });
  txt(s, '04. Write Title Here', [12.196, 7.087, 4.822, 0.505], { ...H2 });
  txt(s, LOREM1, [12.196, 7.874, 5.398, 0.967], { ...BODY });
  box(s, [2.218, 6.581, 6.894, 2.765], { line: { color: CREAM, width: 3 } });
  txt(s, '03. Write Title Here', [3.18, 7.087, 4.822, 0.505], { ...H2 });
  txt(s, LOREM1, [3.18, 7.874, 5.398, 0.967], { ...BODY });
  txt(s, ' Our Amazing Service', [5.606, 1.449, 8.788, 0.909], { ...H1, align: 'center' });
  box(s, [2.218, 3.28, 0.429, 2.765], { fill: WHITE });
  brush(s, 2.401, 5.11, 0.031, 0.046, 0.613, OLIVE, { swap: true });
  brush(s, 2.385, 3.602, 0.031, 0.047, 0.613, OLIVE, { swap: true });
}

function slide16(pptx) {
  const s = pptx.addSlide();
  txt(s, 'The Products We Use', [1.653, 9.208, 8.347, 0.909], { ...H1 });
  txt(s, 'OUR SERVICE', [1.653, 8.606, 3.414, 0.404], { ...KICKER });
  box(s, [17.268, 8.859, 1.078, 1.004], { line: { color: OLIVE, width: 1.5 } });
  glyph(s, [17.6, 9.154, 0.415, 0.415], 'plus', OLIVE);
  box(s, [14.366, 8.859, 1.078, 1.004], { fill: OLIVE, line: { color: OLIVE, width: 1 } });
  box(s, [15.817, 8.859, 1.078, 1.004], { fill: OLIVE, line: { color: OLIVE, width: 1 } });
  box(s, [12.914, 8.859, 1.078, 1.004], { fill: OLIVE, line: { color: OLIVE, width: 1 } });
  box(s, [1.653, 1.133, 6.894, 2.765], { fill: OLIVE, shadow: shade(30, 15, SLATE, 0.4) });
  txt(s, 'Write Title Here', [2.616, 1.638, 4.822, 0.505], { ...H2, color: WHITE });
  txt(s, LOREM1, [2.616, 2.426, 5.398, 0.967], { ...BODY, color: MIST });
  box(s, [1.653, 4.782, 6.894, 2.765], { line: { color: CREAM, width: 3 } });
  txt(s, 'Write Title Here', [2.616, 5.287, 4.822, 0.505], { ...H2 });
  txt(s, LOREM1, [2.616, 6.075, 5.398, 0.967], { ...BODY });
  box(s, [1.653, 1.133, 0.429, 2.765], { fill: WHITE });
  brush(s, 1.836, 2.963, 0.031, 0.047, 0.613, OLIVE, { swap: true });
  brush(s, 1.821, 1.455, 0.031, 0.047, 0.613, OLIVE, { swap: true });
  icon(s, [13.203, 9.111, 0.5, 0.5], WHITE);
  icon(s, [14.652, 9.108, 0.506, 0.506], WHITE);
  icon(s, [16.054, 9.059, 0.605, 0.605], WHITE);
  photo(s, [11.249, 1.133, 3.347, 6.414]);
  photo(s, [15, 1.133, 3.347, 6.414], { fill: GREY });
}

function slide17(pptx) {
  const s = pptx.addSlide();
  box(s, [1.078, 0.881, 17.843, 9.487], { line: { color: OLIVE, width: 3 } });
  box(s, [0, 0, 20, 11.25], { line: { color: OLIVE, width: 20 } });
  brush(s, 16.063, 10.054, 0.63, 0, 1.484, WHITE, { axis: 'x', rows: [0, 2, 4] });
  brush(s, 0.943, 1.595, 0.375, 0, 1.484, WHITE, { rows: [0, 2, 4] });
  photo(s, [2.317, 1.592, 3.965, 4.45]);
  photo(s, [8.018, 1.592, 3.965, 4.45]);
  photo(s, [13.718, 1.592, 3.965, 4.45]);
  txt(s, '01. Write Title Here', [2.317, 6.872, 3.965, 0.505], { ...H2 });
  txt(s, LOREM3, [2.317, 7.659, 3.965, 0.967], { ...BODY });
  txt(s, '$245', [2.317, 8.952, 2.144, 0.707], { ...PRICE });
  txt(s, '02. Write Title Here', [8.018, 6.872, 3.965, 0.505], { ...H2 });
  txt(s, LOREM3, [8.018, 7.659, 3.965, 0.967], { ...BODY });
  txt(s, '$127', [8.018, 8.952, 2.144, 0.707], { ...PRICE });
  txt(s, '03. Write Title Here', [13.718, 6.872, 3.965, 0.505], { ...H2 });
  txt(s, LOREM3, [13.718, 7.659, 3.965, 0.967], { ...BODY });
  txt(s, '$40', [13.718, 8.952, 2.144, 0.707], { ...PRICE });
  box(s, [3.325, 2.909, 1.949, 1.816], { line: { color: WHITE, width: 2.25 } });
  glyph(s, [3.925, 3.442, 0.75, 0.75], 'plus', WHITE);
}

function slide18(pptx) {
  const s = pptx.addSlide();
  rule(s, 0, 1.5, 20, 0, OLIVE, 2.25);
  box(s, [17.344, 10.581, 1.802, 0.083], { fill: OLIVE });
  box(s, [17.344, 10.35, 1.802, 0.087], { fill: OLIVE });
  box(s, [14.641, 10.697, 1.802, 0.087], { fill: OLIVE });
  box(s, [14.641, 10.462, 1.802, 0.09], { fill: OLIVE });
  box(s, [14.641, 10.217, 1.802, 0.087], { fill: OLIVE });
  box(s, [0, 2.668, 20, 7.082], { fill: OLIVE });
  photo(s, [0, 2.668, 20, 7.082]);
  navbar(s, 0);
  txt(s, 'www.hoschen.kitchen.com', [1.166, 10.281, 7.044, 0.438], { ...H3 });
  txt(s, 'Housewares & Home Kitchen Presentation', [1.578, 7.043, 7.044, 0.505], { ...H2, color: WHITE });
  txt(s, 'The people recognize themselves in their commodities; they find their soul in their automobile, hi-fi set, split-level home, kitchen equipment.', [12.266, 5.18, 6.568, 2.073], { ...BODY, fontSize: 20, color: WHITE, bold: true });
  txt(s, 'Breakfast', [1.578, 4.885, 8.83, 1.952], { ...BTN, fontSize: 110, align: 'left' });
}

function slide19(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 0, 1.267, 3.31], { fill: OLIVE });
  box(s, [0, 0, 1.267, 3.286], { fill: OLIVE });
  box(s, [18.749, 7.94, 1.251, 3.31], { fill: OLIVE });
  box(s, [0, 7.94, 10.912, 3.31], { fill: OLIVE });
  rule(s, 0.609, -0.115, 0, 2.235, WHITE, 2.25);
  rule(s, 0.282, -0.026, 0, 1.517, WHITE, 2.25);
  rule(s, 0.936, -0.064, 0, 1.517, WHITE, 2.25);
  rule(s, 8.889, 9.595, 2.235, 0, WHITE, 2.25);
  rule(s, 9.518, 9.921, 1.517, 0, WHITE, 2.25);
  rule(s, 9.556, 9.268, 1.517, 0, WHITE, 2.25);
  txt(s, 'Your Inspiring Kitchen &  Cooking Tools', [10.989, 1.787, 7.548, 1.717], { ...H1 });
  txt(s, 'OUR PORTFOLIO', [10.989, 1.184, 3.414, 0.404], { ...KICKER });
  txt(s, LOREM7, [1.463, 5.289, 7.624, 1.421], { ...BODY });
  txt(s, '01. Write Title Here', [1.463, 4.539, 4.822, 0.505], { ...H2 });
  txt(s, LOREM7, [10.912, 5.289, 7.624, 1.421], { ...BODY });
  txt(s, '02. Write Title Here', [10.912, 4.539, 4.822, 0.505], { ...H2 });
  txt(s, 'Hoschen', [1.463, 8.619, 7.624, 1.784], { ...BTN, fontSize: 100, align: 'left' });
  photo(s, [1.463, 0, 7.624, 3.31]);
  photo(s, [10.912, 7.94, 7.624, 3.31]);
}

function slide20(pptx) {
  const s = pptx.addSlide();
  box(s, [15.333, 0, 4.667, 11.25], { fill: OLIVE });
  box(s, [13.794, 5.987, 4.873, 4.002], { fill: CREAM, line: { color: WHITE, width: 3 } });
  txt(s, LOREM8, [8.575, 7.177, 3.886, 1.421], { ...BODY });
  txt(s, 'Write Title Here', [8.575, 6.474, 3.886, 0.505], { ...H2 });
  txt(s, '$140', [8.575, 8.796, 2.144, 0.707], { ...PRICE });
  txt(s, LOREM8, [8.575, 2.45, 3.886, 1.421], { ...BODY });
  txt(s, 'Write Title Here', [8.575, 1.747, 3.886, 0.505], { ...H2 });
  txt(s, '$120', [8.575, 4.07, 2.144, 0.707], { ...PRICE });
  txt(s, 'Best In The Market', [1.333, 2.679, 4.435, 1.717], { ...H1 });
  txt(s, 'PORTFOLIO', [1.333, 2.076, 3.414, 0.404], { ...KICKER });
  txt(s, LOREM6, [1.333, 4.856, 5.46, 2.784], { ...BODY });
  box(s, [1.333, 8.17, 3.154, 1.004], { fill: OLIVE, line: { color: NAVY, width: 1 } });
  txt(s, 'Add to Cart', [2.195, 8.453, 1.997, 0.438], { ...BTN });
  glyph(s, [1.629, 8.476, 0.392, 0.392], 'cart', WHITE);
  photo(s, [13.794, 1.261, 4.873, 4.002], { line: { color: WHITE, width: 3 } });
  photo(s, [14.284, 6.39, 3.893, 3.197], { fill: GREY });
  box(s, [15.255, 2.354, 1.949, 1.816], { line: { color: WHITE, width: 2.25 } });
  glyph(s, [15.855, 2.887, 0.75, 0.75], 'plus', WHITE);
}

function slide21(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 6.354, 20, 4.896], { fill: OLIVE });
  box(s, [0.667, 10.305, 19.333, 0.342], { fill: WHITE });
  box(s, [0.667, 6.354, 0.328, 4.293], { fill: WHITE });
  brush(s, 7.546, 10.44, 0.031, 0.048, 1.087, OLIVE, { axis: 'x', flip: true });
  brush(s, 4.841, 10.44, 0.031, 0.048, 1.087, OLIVE, { axis: 'x', flip: true });
  brush(s, 2.136, 10.44, 0.031, 0.048, 1.087, OLIVE, { axis: 'x', flip: true });
  brush(s, 15.443, 10.432, 0.031, 0.048, 1.087, OLIVE, { axis: 'x', flip: true });
  brush(s, 12.738, 10.432, 0.031, 0.048, 1.087, OLIVE, { axis: 'x', flip: true });
  brush(s, 10.034, 10.432, 0.031, 0.048, 1.087, OLIVE, { axis: 'x', flip: true });
  brush(s, 17.71, 10.432, 0.031, 0.048, 1.087, OLIVE, { axis: 'x', flip: true });
  brush(s, 0.794, 6.589, 0.031, 0.046, 1.087, OLIVE, { swap: true });
  brush(s, 0.794, 9.14, 0.031, 0.046, 1.087, OLIVE, { swap: true });
  txt(s, LOREM9, [5.577, 7.729, 3.886, 0.967], { ...BODY, color: WHITE });
  txt(s, 'Write Title Here', [5.577, 7.026, 3.886, 0.505], { ...H2, color: WHITE });
  txt(s, '$135', [5.577, 8.894, 2.144, 0.707], { ...PRICE, color: WHITE });
  txt(s, LOREM9, [14.961, 7.729, 3.886, 0.967], { ...BODY, color: WHITE });
  txt(s, 'Write Title Here', [14.961, 7.026, 3.886, 0.505], { ...H2, color: WHITE });
  txt(s, '$102', [14.961, 8.894, 2.144, 0.707], { ...PRICE, color: WHITE });
  txt(s, 'HOSCHEN', [17.105, 0.461, 2.286, 0.505], { fontFace: POPM, fontSize: 24, color: OLIVE, bold: true, align: 'right' });
  txt(s, 'Best Seller Product', [1.134, 1.779, 14.943, 0.909], { ...H1 });
  txt(s, 'PORTFOLIO', [1.134, 1.177, 3.414, 0.404], { ...KICKER });
  txt(s, 'Lorem ipsum dolor sit amet, elit, sed do ni labore et. aliqua. Nibh posuere morbi leo urna molestie at elementum eu. feugiat pretium nibh this amet commodo nulla facilisi nullam vehicula ipsum a. Sem vitae justo eget magna fermentum nisl ullamcorper magna tempor incididunT. Fermentum odio eu feugiat pretium nibh ipsum consequat nisl ipsum ac placerat vestibulum lectus.', [1.134, 3.802, 17.731, 1.421], { ...BODY });
  txt(s, 'Elegant & Engineered', [1.918, 2.915, 5.506, 0.37], { ...H2, fontSize: 16 });
  icon(s, [1.134, 2.85, 0.5, 0.5], OLIVE);
  photo(s, [1.57, 6.896, 2.959, 2.835]);
  photo(s, [10.954, 6.896, 2.959, 2.835]);
}

function slide22(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 3.163, 20, 4.923], { fill: OLIVE });
  rule(s, 0, 5.625, 2.262, 0, WHITE, 3);
  rule(s, 0, 6.239, 1.611, 0, WHITE, 3);
  rule(s, 0.786, 3.586, 0, 2.85, WHITE, 3);
  rule(s, 17.738, 5.625, 2.262, 0, WHITE, 3);
  rule(s, 18.389, 6.239, 1.611, 0, WHITE, 3);
  rule(s, 19.214, 3.586, 0, 2.85, WHITE, 3);
  box(s, [1.385, 1.281, 17.231, 8.688], { fill: WHITE, shadow: shade(64, 40, SLATE, 0.4) });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur.', [6.339, 3.423, 3.199, 0.967], { ...BODY });
  txt(s, 'Write Title Here', [6.339, 4.631, 3.199, 0.505], { ...H2 });
  txt(s, '$130', [6.339, 2.515, 2.144, 0.707], { ...PRICE });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur.', [6.339, 7.022, 3.199, 0.967], { ...BODY });
  txt(s, 'Write Title Here', [6.339, 8.23, 3.199, 0.505], { ...H2 });
  txt(s, '$150', [6.339, 6.114, 2.144, 0.707], { ...PRICE });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur.', [10.462, 3.423, 3.199, 0.967], { ...BODY, align: 'right' });
  txt(s, 'Write Title Here', [10.462, 4.631, 3.199, 0.505], { ...H2, align: 'right' });
  txt(s, '$80', [11.517, 2.515, 2.144, 0.707], { ...PRICE, align: 'right' });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur.', [10.462, 7.022, 3.199, 0.967], { ...BODY, align: 'right' });
  txt(s, 'Write Title Here', [10.462, 8.23, 3.199, 0.505], { ...H2, align: 'right' });
  txt(s, '$120', [11.517, 6.114, 2.144, 0.707], { ...PRICE, align: 'right' });
  photo(s, [2.769, 2.515, 2.446, 2.621]);
  photo(s, [2.769, 6.114, 2.446, 2.621]);
  photo(s, [14.784, 2.515, 2.446, 2.621]);
  photo(s, [14.784, 6.114, 2.446, 2.589]);
}

function slide23(pptx) {
  const s = pptx.addSlide();
  rule(s, 0, 1.5, 20, 0, OLIVE, 2.25);
  navbar(s, 2);
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. this aliqua. Nibh posuere placerat vestibulum lectus mauris odio eu feugiat pretium.', [1.166, 3.345, 8.966, 1.421], { ...BODY });
  txt(s, '1200+', [1.166, 2.328, 2.144, 0.707], { ...PRICE });
  txt(s, 'Shop Your Cllection', [5.583, 2.429, 3.886, 0.505], { ...H2, align: 'right' });
  txt(s, 'The Portfolio', [12.798, 2.725, 6.036, 0.909], { ...H1, align: 'right' });
  txt(s, '100% Top In The Market', [13.328, 3.999, 5.506, 0.37], { ...H2, fontSize: 16, align: 'right' });
  photo(s, [0, 6.473, 3.692, 4.796]);
  photo(s, [4.49, 7.019, 2.852, 3.704]);
  photo(s, [8.154, 6.48, 3.692, 4.79]);
  photo(s, [12.644, 7.019, 2.852, 3.704]);
  photo(s, [16.308, 6.473, 3.692, 4.796]);
}

function slide24(pptx) {
  const s = pptx.addSlide();
  box(s, [-0.286, 7.676, 2.257, 0.132], { fill: OLIVE });
  box(s, [-0.286, 7.318, 2.257, 0.137], { fill: OLIVE });
  box(s, [-0.286, 6.944, 2.257, 0.132], { fill: OLIVE });
  box(s, [-0.286, 8.878, 2.257, 0.127], { fill: { color: SMOKE, transparency: 50 }, flipH: true });
  box(s, [-0.286, 8.526, 2.257, 0.132], { fill: { color: SMOKE, transparency: 50 }, flipH: true });
  box(s, [8.103, 6.355, 11.897, 3.239], { fill: OLIVE });
  device(s, [1.409, 5.982, 7.067, 3.985], 'phone');
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et. commodo nulla facilisi nullam vehicula ipsum a. Sem integer vitae justo eget magna fermentum. odio eu feugiat pretium nibh ipsum.', [10.903, 7.037, 7.688, 1.876], { ...BODY, color: WHITE });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. aliqua. Nibh tortor id aliquet lectus. Amet nisl purus in mollis nunc sed id. Posuere morbi leo urna molestie at elementum eu pulvinar mattis nunc sed blandit. Mauris commodo quis imperdiet a massa lacus non odio, nulla facilisi nullam.  ', [10.903, 1.442, 7.688, 2.784], { ...BODY });
  txt(s, 'Download Now', [10.903, 4.505, 3.404, 0.438], { ...H3, fontFace: POP, bold: true });
  txt(s, 'Download Our App', [1.408, 2.528, 7.067, 0.909], { ...H1 });
  txt(s, '100% Fast Delivery After Checkout', [1.408, 3.666, 4.387, 0.37], { ...BODY16, lineSpacingMultiple: 1 });
  txt(s, 'MOCKUP DEVICE', [1.408, 1.946, 3.414, 0.404], { ...KICKER });
  photo(s, [1.58, 6.192, 6.724, 3.566]);
}

function slide25(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 0, 1.457, 11.25], { fill: OLIVE });
  rule(s, 0.729, -0.013, 0, 2.235, WHITE, 2.25);
  rule(s, 0.414, -0.01, 0, 1.517, WHITE, 2.25);
  rule(s, 1.043, -0.01, 0, 1.517, WHITE, 2.25);
  rule(s, 0.729, 9.015, 0, 2.235, WHITE, 2.25);
  rule(s, 0.414, 9.73, 0, 1.517, WHITE, 2.25);
  rule(s, 1.043, 9.73, 0, 1.517, WHITE, 2.25);
  device(s, [12.181, 2.781, 9.462, 5.477], 'laptop');
  txt(s, 'HOSCHEN', [0.333, 3.277, 0.791, 4.695], { ...BTN, fontSize: 24, bold: true });
  txt(s, 'A Website Makes Shopping Easier', [3.088, 2.428, 6.641, 1.717], { ...H1 });
  txt(s, 'MOCKUP DEVICE', [3.088, 1.825, 3.414, 0.404], { ...KICKER });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur elit, sed do nisl eiusmod tempor incididunt ut labore et. this aliqua. Nibh posuere morbi leo urna molestie at nibh in ipsum nisl ullamcorper magna ac placerat lectus mauris adipiscing risus a diam.', [3.088, 4.605, 6.912, 2.33], { ...BODY });
  icon(s, [3.088, 7.551, 0.707, 0.707], OLIVE);
  txt(s, 'Modern Design', [4.384, 7.652, 5.345, 0.505], { ...H2 });
  txt(s, [
    { text: LOREM10, options: { ...BODY, bullet: { characterCode: '2751', indent: 22.5 }, breakLine: true } },
    { text: LOREM10, options: { ...BODY, bullet: { characterCode: '2751', indent: 22.5 } } },
  ], [3.088, 8.458, 6.641, 0.967]);
  photo(s, [13.236, 3.081, 7.447, 4.73]);
}

function slide26(pptx) {
  const s = pptx.addSlide();
  box(s, [0, 0, 20, 11.25], { line: { color: OLIVE, width: 20 } });
  brush(s, 19.273, 0.568, 0.727, 0, 1.466, OLIVE, { rows: [1, 3] });
  brush(s, 0, 0.568, 0.727, 0, 1.466, OLIVE, { rows: [1, 3] });
  brush(s, 19.273, 9.868, 0.727, 0, 1.468, OLIVE, { rows: [1, 3] });
  brush(s, 0, 9.868, 0.727, 0, 1.468, OLIVE, { rows: [1, 3] });
  device(s, [1.795, 3.382, 5.401, 4.486], 'monitor');
  txt(s, 'Feedback Should Live On Your Website', [9.416, 1.897, 8.25, 1.717], { ...H1 });
  txt(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do nisl eiusmod tempor incididunt ut labore et. aliqua. Nibh tortor id aliquet lectus. Amet nisl purus in mollis nunc sed id. Posuere morbi leo urna molestie at elementum eu pulvinar mattis nunc sed blandit. Mauris vehicula ipsum a. Sem integer vitae justo consequat nisl et.', [9.416, 4.119, 8.788, 2.33], { ...BODY });
  txt(s, 'MOCKUP DEVICE', [9.422, 1.212, 3.414, 0.404], { ...KICKER });
  box(s, [9.416, 7.273, 6.894, 2.765], { line: { color: CREAM, width: 3 } });
  txt(s, '01. Write Title Here', [10.378, 7.779, 4.822, 0.505], { ...H2 });
  txt(s, LOREM1, [10.378, 8.566, 5.398, 0.967], { ...BODY });
  photo(s, [2.086, 3.641, 4.821, 2.718]);
}

function slide27(pptx) {
  const s = pptx.addSlide();
  txt(s, 'Choose Your Pricing Plan', [4.897, 0.961, 10.206, 0.909], { ...H1, align: 'center' });
  box(s, [1.376, 2.91, 5.147, 7.379], { fill: WHITE, shadow: shade(12, 9, GREY, 0.4) });
  txt(s, '126', [3.259, 4.439, 1.382, 0.924], { ...PRICETAG });
  txt(s, '$', [2.971, 4.553, 0.747, 0.558], { ...PRICETAG, fontSize: 20, color: OLIVE });
  txt(s, '/Years', [4.35, 4.901, 0.911, 0.421], { ...BODY, fontSize: 14, color: OLIVE });
  txt(s, 'Elegant & Engineered', [2.414, 5.742, 3.072, 0.512], { ...PLAN });
  txt(s, 'Modern Design', [2.414, 7.012, 3.072, 0.512], { ...PLAN });
  txt(s, 'Minimalist Kitchen', [2.414, 6.377, 3.072, 0.512], { ...PLAN });
  txt(s, 'Top In The Market', [2.414, 7.647, 3.072, 0.512], { ...PLAN });
  box(s, [2.561, 8.719, 2.776, 0.957], { fill: OLIVE });
  box(s, [2.557, 2.923, 2.785, 1.09], { fill: OLIVE });
  box(s, [7.426, 2.91, 5.147, 7.379], { fill: WHITE, shadow: shade(12, 9, GREY, 0.4) });
  txt(s, '390', [9.309, 4.439, 1.382, 0.924], { ...PRICETAG });
  txt(s, '$', [9.022, 4.553, 0.747, 0.558], { ...PRICETAG, fontSize: 20, color: OLIVE });
  txt(s, '/Years', [10.401, 4.901, 0.911, 0.421], { ...BODY, fontSize: 14, color: OLIVE });
  box(s, [13.477, 2.91, 5.147, 7.379], { fill: WHITE, shadow: shade(12, 9, GREY, 0.4) });
  txt(s, '589', [15.36, 4.439, 1.382, 0.924], { ...PRICETAG });
  txt(s, '$', [15.072, 4.553, 0.747, 0.558], { ...PRICETAG, fontSize: 20, color: OLIVE });
  txt(s, '/Years', [16.451, 4.901, 0.911, 0.421], { ...BODY, fontSize: 14, color: OLIVE });
  txt(s, 'Basic', [2.791, 3.249, 2.316, 0.438], { ...BTN });
  txt(s, 'Standart Plan', [8.842, 3.249, 2.316, 0.438], { ...BTN });
  box(s, [8.612, 8.719, 2.776, 0.957], { fill: OLIVE });
  box(s, [8.607, 2.923, 2.785, 1.09], { fill: OLIVE });
  txt(s, 'Standart', [8.842, 3.249, 2.316, 0.438], { ...BTN });
  txt(s, 'Standart Plan', [15.081, 3.236, 2.316, 0.438], { ...BTN });
  box(s, [14.658, 2.91, 2.785, 1.09], { fill: OLIVE });
  txt(s, 'Advanced', [14.893, 3.236, 2.316, 0.438], { ...BTN });
  box(s, [14.662, 8.719, 2.776, 0.957], { fill: OLIVE });
  txt(s, 'Buy Now', [2.791, 8.978, 2.316, 0.438], { ...BTN });
  txt(s, 'Buy Now', [8.842, 8.978, 2.316, 0.438], { ...BTN });
  txt(s, 'Buy Now', [14.893, 8.978, 2.316, 0.438], { ...BTN });
  txt(s, 'Elegant & Engineered', [8.464, 5.742, 3.072, 0.512], { ...PLAN });
  txt(s, 'Modern Design', [8.464, 7.012, 3.072, 0.512], { ...PLAN });
  txt(s, 'Minimalist Kitchen', [8.464, 6.377, 3.072, 0.512], { ...PLAN });
  txt(s, 'Top In The Market', [8.464, 7.647, 3.072, 0.512], { ...PLAN });
  txt(s, 'Elegant & Engineered', [14.515, 5.746, 3.072, 0.512], { ...PLAN });
  txt(s, 'Modern Design', [14.515, 7.016, 3.072, 0.512], { ...PLAN });
  txt(s, 'Minimalist Kitchen', [14.515, 6.381, 3.072, 0.512], { ...PLAN });
  txt(s, 'Top In The Market', [14.515, 7.65, 3.072, 0.512], { ...PLAN });
}

function slide28(pptx) {
  const s = pptx.addSlide();
  box(s, [10, 0, 10, 11.25], { fill: MIST });
  box(s, [1.409, 1.112, 12.924, 9.026], { fill: OLIVE, shadow: shade(30, 15, SLATE, 0.4) });
  box(s, [1.409, 2.892, 0.847, 5.465], { fill: WHITE });
  box(s, [14.331, 1.112, 4.259, 9.026], { fill: WHITE, shadow: shade(30, 15, SLATE, 0.4) });
  box(s, [17.631, 10.705, 1.802, 0.083], { fill: OLIVE });
  box(s, [17.631, 10.474, 1.802, 0.087], { fill: OLIVE });
  box(s, [0.567, 0.693, 1.802, 0.083], { fill: OLIVE });
  box(s, [0.567, 0.461, 1.802, 0.087], { fill: OLIVE });
  txt(s, '“At home in Ghaziabad, everyone is a pure vegetarian. In fact, when I want to cook non-veg there, my mum shoos me out on the terrace where I have my cooking utensils. I\'m told categorically that whatever non-veg or egg, etc., that I have to cook, I should do upstairs and not enter her kitchen at all.”', [3.43, 3.151, 8.883, 4.982], { ...PLAN, fontSize: 28, color: WHITE });
  txt(s, 'SURESH RAINA', [4.349, 2.389, 7.044, 0.438], { ...BTN, fontFace: ASI, bold: true });
  txt(s, '#QUETES', [6.164, 8.457, 3.414, 0.404], { ...KICKER, color: MIST, align: 'center' });
  txt(s, 'HOSCHEN', [1.511, 3.355, 0.644, 4.541], { ...H1, fontSize: 18, bold: true, align: 'center' });
  photo(s, [14.331, 1.112, 4.049, 9.026]);
}

function slide29(pptx) {
  const s = pptx.addSlide();
  box(s, [19.168, 8.023, 0.832, 1.741], { fill: OLIVE });
  txt(s, 'Let\'s Chat, Contact Our Address', [10.835, 2.127, 7.548, 1.717], { ...H1, align: 'right' });
  txt(s, 'CONTACT US', [14.968, 1.542, 3.414, 0.404], { ...KICKER, align: 'right' });
  glyph(s, [14.579, 3.934, 3.804, 2.482], 'uturn', OLIVE);
  txt(s, 'Our Address', [2.105, 8.023, 3.046, 0.505], { ...H2 });
  txt(s, [
    { text: '175 West Valley Freeway, California,', options: { ...BODY, breakLine: true } },
    { text: 'United States', options: { ...BODY } },
  ], [2.105, 8.797, 4.815, 0.967]);
  txt(s, 'Email', [8.994, 8.023, 2.212, 0.505], { ...H2 });
  txt(s, [
    { text: 'info@hoschen.com', options: { ...BODY, breakLine: true } },
    { text: 'support@hoschen.com', options: { ...BODY } },
  ], [8.994, 8.797, 3.516, 0.967]);
  txt(s, 'Telephone', [14.584, 8.023, 2.51, 0.505], { ...H2 });
  txt(s, [
    { text: '(+021) 786 5632', options: { ...BODY, breakLine: true } },
    { text: '(+123) 786 5634', options: { ...BODY } },
  ], [14.584, 8.797, 3.311, 0.967]);
  photo(s, [0, 1.486, 8.994, 4.986]);
}

function slide30(pptx) {
  const s = pptx.addSlide();
  rule(s, 0, 1.5, 20, 0, OLIVE, 2.25);
  box(s, [1.451, 2.872, 17.083, 8.378], { fill: OLIVE, line: { color: NAVY, width: 1 } });
  photo(s, [1.451, 2.872, 17.098, 8.378]);
  navbar(s, 0);
  txt(s, 'Thank You', [4.562, 7.076, 10.877, 1.952], { ...BTN, fontSize: 110 });
  txt(s, 'For Watching My Presentation Template', [6.478, 9.097, 7.044, 0.505], { ...H2, color: WHITE, align: 'center' });
  txt(s, '“Cooking is one of the strongest ceremonies for life. When recipes are put together, the kitchen is a chemical laboratory involving air, fire, water and the earth. This is what gives value to humans and elevates their spiritual qualities. If you take a frozen box and stick it in the microwave, you become connected to the factory.”', [5.115, 4.52, 9.769, 2.33], { ...PLAN, color: WHITE });
}


// ---------------------------------------------------------------- assemble
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'HOSCHEN', width: 20, height: 11.25 });
  pptx.layout = 'HOSCHEN';
  pptx.author = 'Hoschen';
  pptx.title = 'Hoschen Kitchen Equipment Presentation';

  const slides = [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
    slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
  ];
  slides.forEach(fn => fn(pptx));

  const out = path.join(__dirname, path.basename(__filename, '.js') + '.pptx');
  return pptx.writeFile({ fileName: out }).then(() => console.log('wrote', out));
}

build().catch(err => { console.error(err); process.exit(1); });
