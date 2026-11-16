/**
 * "Holiday Co." travel deck - 30 slides, 10 x 7.5 in.
 * Rebuilt from scratch with pptxgenjs; raster mockups are drawn as native placeholders.
 *
 *   node 0bbba078-e6be-449e-b188-b918ff1eb452_grok_final.js
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- theme ----
const BLUE = '056AFF'; // accent1
const BLUE2 = '078EFD'; // accent2
const BLUE3 = '05A0FF'; // accent3
const BLUE4 = '11B8F3'; // accent4
const BLUE5 = '39C4F5'; // accent5
const WHITE = 'FFFFFF';
const BLACK = '000000';
const GRAY = 'D9D9D9'; // white @ 85% luminance - rules, tracks, master chrome text
const MAPGRAY = 'D9D9D9';
const PLACEHOLDER = 'F2F2F2'; // stand-in screen fill for the deck's raster mockups
const FRAME = '1A1A1A'; // device-frame body of those mockups

const HEAD = 'Roboto Black'; // +mj-lt
const BODY = 'Raleway'; // +mn-lt

// Dark scrim: black at 85% opacity, exactly as the reference overlays it.
const SCRIM = { color: BLACK, transparency: 15 };
// Fresh object per call - pptxgenjs rewrites shadow props in place.
function cardShadow() {
  return { type: 'outer', color: BLACK, opacity: 0.2, blur: 23, offset: 0, angle: 90 };
}

// ------------------------------------------------------------ lorem text ----
const L_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, ';
const L_MED = L_SHORT + 'magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. ';
const L_LONG = L_MED + 'Nunc viverra imperdiet enim. ';
const L_VIVERRA = L_MED + 'Nunc viverra';
const L_TELLUS = L_MED + 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque';
const L_CONGUE = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere congue';
const L_PORTTITOR = L_CONGUE + ' porttitor congue.';
const L_TINY = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue';
const L_AMET = L_SHORT + 'magna sed pulvinar ultricies, purus lectus malesuada libero, sit ' + 'amet';
const L_PULVINAR = L_SHORT + 'magna sed pulvinar';
const L_MAGNA = L_SHORT + 'magna ';
const L_MALESUADA = L_SHORT + 'magna sed pulvinar ultricies, purus lectus malesuada';
const L_LIBERO = L_SHORT + 'magna sed pulvinar ultricies, purus lectus malesuada libero';
const L_EROS = L_SHORT + 'magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros';
const L_ELIT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ';

// ------------------------------------------------------------- primitives ---
/** Plain filled rectangle. */
function rect(s, x, y, w, h, fill) {
  s.addShape('rect', { x, y, w, h, fill, line: { type: 'none' } });
}

/** Rounded rectangle; r is the corner radius in inches. */
function round(s, x, y, w, h, fill, r, extra) {
  s.addShape('roundRect', Object.assign({ x, y, w, h, fill, line: { type: 'none' }, rectRadius: r }, extra || {}));
}

/** Fully rounded "pill" (adj 50000 in the source deck). */
function pill(s, x, y, w, h, fill, line) {
  s.addShape('roundRect', {
    x, y, w, h, rectRadius: Math.min(w, h) / 2,
    fill: fill || { type: 'none' },
    line: line || { type: 'none' },
  });
}

/** White content card with the deck's soft drop shadow (adj 8586). */
function card(s, x, y, w, h, radius) {
  s.addShape('roundRect', {
    x, y, w, h,
    rectRadius: radius === undefined ? 0.0859 * Math.min(w, h) : radius,
    fill: { color: WHITE }, line: { type: 'none' }, shadow: cardShadow(),
  });
}

/** Text box. `o` accepts every pptxgenjs text option. */
function text(s, body, o) {
  s.addText(body, Object.assign({ fontFace: BODY, fontSize: 12, color: BLACK, valign: 'middle' }, o));
}

/** Body copy: 9 pt, 1.5 line spacing - the deck's paragraph default. */
function para(s, body, x, y, w, h, o) {
  text(s, body, Object.assign({ x, y, w, h, fontSize: 9, lineSpacingMultiple: 1.5 }, o || {}));
}

/** Section/card heading: 12 pt regular. */
function label(s, body, x, y, w, h, o) {
  text(s, body, Object.assign({ x, y, w, h, fontSize: 12 }, o || {}));
}

/** Big accent figure: 21 pt Roboto Black in accent blue. */
function stat(s, value, x, y, w, o) {
  text(s, value, Object.assign({ x, y, w, h: 0.454, fontSize: 21, fontFace: HEAD, color: BLUE }, o || {}));
}

/** Title placeholder: coloured runs, Roboto Black, 90% line spacing. */
function title(s, runs, x, y, w, h, size, base, align) {
  s.addText(
    runs.map(function (r) { return { text: r[0], options: { color: r[1] || base } }; }),
    {
      x, y, w, h, fontFace: HEAD, fontSize: size, color: base,
      align: align || 'left', valign: 'middle', lineSpacingMultiple: 0.9,
    }
  );
}

/** Thin vertical divider between two stats. */
function divider(s, x, y, h) {
  s.addShape('line', { x, y, w: 0, h, line: { color: GRAY, width: 1 } });
}

/** Skill/progress bar: grey track with a blue fill and a label + percentage. */
function progress(s, name, pct, x, y, trackW, fillW, labelW, pctX) {
  label(s, name, x, y - 0.291, labelW, 0.303);
  label(s, pct, pctX, y - 0.291, 0.95, 0.303, { align: 'right' });
  pill(s, x, y, trackW, 0.121, { color: GRAY });
  pill(s, x, y, fillW, 0.121, { color: BLUE });
}

// ------------------------------------------------------------------ icons ---
// Each glyph is a list of primitives inside a 0..1 box: [shape, x, y, w, h, 'fg'|'bg', radius?]
const ICONS = {
  car: [
    ['roundRect', 0.22, 0.00, 0.56, 0.52, 'fg', 0.22],
    ['rect', 0.46, 0.02, 0.07, 0.34, 'bg'],
    ['roundRect', 0.00, 0.30, 1.00, 0.55, 'fg', 0.25],
    ['roundRect', 0.09, 0.44, 0.82, 0.11, 'bg', 0.05],
    ['ellipse', 0.17, 0.57, 0.17, 0.24, 'bg'],
    ['ellipse', 0.66, 0.57, 0.17, 0.24, 'bg'],
  ],
  bed: [
    ['rect', 0.04, 0.00, 0.12, 0.58, 'fg'],
    ['rect', 0.84, 0.00, 0.12, 0.58, 'fg'],
    ['rect', 0.04, 0.14, 0.92, 0.13, 'fg'],
    ['rect', 0.46, 0.14, 0.08, 0.16, 'bg'],
    ['roundRect', 0.00, 0.40, 1.00, 0.40, 'fg', 0.14],
    ['rect', 0.00, 0.78, 1.00, 0.14, 'fg'],
    ['rect', 0.04, 0.86, 0.10, 0.14, 'fg'],
    ['rect', 0.86, 0.86, 0.10, 0.14, 'fg'],
  ],
  food: [ // burger: domed bun, lettuce, patty, base bun
    ['ellipse', 0.06, 0.00, 0.88, 0.52, 'fg'],
    ['rect', 0.06, 0.24, 0.88, 0.12, 'fg'],
    ['roundRect', 0.00, 0.30, 1.00, 0.22, 'fg', 0.09],
    ['rect', 0.03, 0.28, 0.94, 0.04, 'bg'],
    ['roundRect', 0.00, 0.50, 1.00, 0.18, 'fg', 0.07],
    ['rect', 0.03, 0.48, 0.94, 0.04, 'bg'],
    ['roundRect', 0.03, 0.66, 0.94, 0.34, 'fg', 0.12],
    ['rect', 0.05, 0.64, 0.90, 0.04, 'bg'],
  ],
  beach: [
    ['ellipse', 0.06, 0.02, 0.88, 0.60, 'fg'],
    ['rect', 0.00, 0.32, 1.00, 0.68, 'bg'],
    ['rect', 0.06, 0.24, 0.88, 0.10, 'fg'],
    ['rect', 0.44, 0.00, 0.09, 0.10, 'fg'],
    ['rect', 0.46, 0.28, 0.07, 0.50, 'fg'],
    ['rtTriangle', 0.16, 0.62, 0.84, 0.24, 'fg'],
    ['rect', 0.00, 0.86, 1.00, 0.12, 'fg'],
  ],
  laptop: [
    ['rect', 0.10, 0.00, 0.80, 0.62, 'fg'],
    ['rect', 0.20, 0.10, 0.60, 0.42, 'bg'],
    ['rect', 0.00, 0.62, 1.00, 0.20, 'fg'],
    ['rect', 0.00, 0.82, 1.00, 0.18, 'fg'],
    ['ellipse', 0.26, 0.60, 0.16, 0.22, 'bg'],
    ['ellipse', 0.58, 0.60, 0.16, 0.22, 'bg'],
  ],
  check: [
    ['ellipse', 0.00, 0.00, 1.00, 1.00, 'fg'],
    ['ellipse', 0.14, 0.14, 0.72, 0.72, 'bg'],
  ],
  pin: [
    ['ellipse', 0.00, 0.00, 1.00, 0.83, 'fg'],
    ['triangle', 0.24, 0.52, 0.52, 0.48, 'fg', 0, 180],
    ['ellipse', 0.29, 0.22, 0.42, 0.42, 'bg'],
  ],
};

/** Draw a glyph from ICONS scaled into the given box. */
function icon(s, name, x, y, w, h, fg, bg) {
  const back = bg || WHITE;
  ICONS[name].forEach(function (p) {
    const opt = {
      x: x + p[1] * w, y: y + p[2] * h, w: p[3] * w, h: p[4] * h,
      fill: { color: p[5] === 'fg' ? fg : back }, line: { type: 'none' },
    };
    if (p[6]) opt.rectRadius = p[6] * Math.min(w, h);
    if (p[7]) opt.rotate = p[7];
    s.addShape(p[0], opt);
  });
  if (name === 'check') {
    s.addShape('custGeom', {
      x, y, w, h, fill: { color: fg }, line: { type: 'none' },
      points: [
        { x: 0.28 * w, y: 0.50 * h, moveTo: true }, { x: 0.43 * w, y: 0.65 * h },
        { x: 0.72 * w, y: 0.33 * h }, { x: 0.80 * w, y: 0.43 * h },
        { x: 0.43 * w, y: 0.82 * h }, { x: 0.20 * w, y: 0.59 * h }, { close: true },
      ],
    });
  }
}

/** Blue rounded tile (adj 33060) with a white glyph inside it. */
function iconTile(s, x, y, name, gx, gy, gw, gh) {
  round(s, x, y, 0.543, 0.543, { color: BLUE }, 0.18);
  icon(s, name, gx, gy, gw, gh, WHITE, BLUE);
}

/** Circular social badge (twitter / facebook / linkedin); the letter is knocked out in white. */
function social(s, x, y, d, letter, color) {
  s.addShape('ellipse', { x, y, w: d, h: d * 0.98, fill: { color }, line: { type: 'none' } });
  text(s, letter, {
    x: x - 0.05, y, w: d + 0.1, h: d * 0.98, align: 'center', valign: 'middle',
    fontSize: Math.round(d * 46), bold: true, color: WHITE,
  });
}

/** The deck's three social badges, laid out left to right. */
function socialRow(s, x, y, d, color) {
  [['t', 0], ['f', 1], ['in', 2]].forEach(function (b) {
    social(s, x + b[1] * 0.2205, y, d, b[0], color);
  });
}

/** Row of five accent stars (testimonial rating). */
function stars(s, x, y) {
  for (let i = 0; i < 5; i++) {
    s.addShape('star5', { x: x + i * 0.2203, y, w: 0.139, h: 0.139, fill: { color: BLUE }, line: { type: 'none' } });
  }
}

/**
 * Stand-ins for the deck's two raster device mockups: a dark frame with a light
 * screen, captioned "[image]". Geometry follows the originals' proportions.
 */
function phoneMockup(s, x, y, w, h) {
  round(s, x, y, w, h, { color: FRAME }, 0.135 * w);
  round(s, x + 0.022 * w, y + 0.056 * h, 0.962 * w, 0.898 * h, { color: PLACEHOLDER }, 0.1 * w);
  text(s, '[image]', { x, y: y + h / 2 - 0.15, w, h: 0.3, align: 'center', fontSize: 9, color: '9A9A9A' });
}

function laptopMockup(s, x, y, w, h) {
  round(s, x + 0.097 * w, y, 0.809 * w, 0.87 * h, { color: FRAME }, 0.02 * w); // lid
  rect(s, x + 0.125 * w, y + 0.058 * h, 0.755 * w, 0.781 * h, { color: PLACEHOLDER }); // screen
  round(s, x, y + 0.87 * h, w, 0.075 * h, { color: 'C9CDD2' }, 0.012 * w); // base
  round(s, x + 0.42 * w, y + 0.895 * h, 0.16 * w, 0.03 * h, { color: '9AA0A6' }, 0.006 * w); // notch
  text(s, '[image]', { x, y: y + h * 0.44 - 0.15, w, h: 0.3, align: 'center', fontSize: 9, color: '9A9A9A' });
}

// ------------------------------------------------------------- page chrome --
/** Chrome inherited from the slide master: light grey, present on every slide. */
function masterChrome(s, pageNo) {
  pill(s, 8.928, 0.215, 0.871, 0.3125, { color: BLUE });
  text(s, 'Travel', { x: 9.082, y: 0.245, w: 0.563, h: 0.252, fontSize: 9, fontFace: HEAD, color: WHITE, align: 'center', wrap: false });
  text(s, 'Holiday Co.', { x: 0.201, y: 0.154, w: 1.075, h: 0.303, fontFace: HEAD, color: GRAY, wrap: false });
  s.addText(
    [{ text: pageNo + ' ', options: { fontSize: 12 } }, { text: '/30', options: { fontSize: 7.5 } }],
    { x: 0.201, y: 7.027, w: 0.595, h: 0.303, fontFace: HEAD, color: GRAY, valign: 'middle', wrap: false }
  );
  [['Facebook', 7.357, 0.688], ['Pinterest', 8.303, 0.631], ['Twitter', 9.265, 0.544]].forEach(function (t) {
    text(s, t[0], { x: t[1], y: 7.086, w: t[2], h: 0.227, fontSize: 7.5, color: GRAY, align: 'center', wrap: false });
  });
}

/**
 * White chrome repeated on individual slides (sits on top of the dark scrims).
 * `p` selects which pieces the slide carries and nudges their exact offsets.
 */
function slideChrome(s, p) {
  if (p.pillY !== undefined) {
    pill(s, p.pillX || 8.928, p.pillY, 0.871, 0.312, { color: BLUE });
    text(s, 'Travel', { x: (p.pillX || 8.928) + 0.154, y: p.pillY + 0.031, w: 0.563, h: 0.252, fontSize: 9, fontFace: HEAD, color: WHITE, align: 'center', wrap: false });
  }
  if (p.logoY !== undefined) {
    text(s, 'Holiday Co.', { x: p.logoX || 0.201, y: p.logoY, w: 1.075, h: 0.303, fontFace: HEAD, color: WHITE, wrap: false });
  }
  if (p.page) {
    s.addText(
      [{ text: p.page + ' ', options: { fontSize: 12 } }, { text: '/30', options: { fontSize: 7.5 } }],
      { x: 0.201, y: p.pageY, w: p.pageW || 0.6, h: 0.303, fontFace: HEAD, color: WHITE, valign: 'middle', wrap: false }
    );
  }
  if (p.socialY !== undefined) {
    [['Facebook', 7.357, 0.688], ['Pinterest', 8.303, 0.631], ['Twitter', 9.265, 0.544]].forEach(function (t) {
      text(s, t[0], { x: t[1], y: p.socialY, w: t[2], h: 0.227, fontSize: 7.5, color: WHITE, align: 'center', wrap: false });
    });
  }
}

/** "SECTION BREAK" kicker + subtitle used by the five divider slides. */
function sectionLabels(s, x, kickerY, subY, w, sub, size, align) {
  text(s, 'SECTION BREAK', { x, y: kickerY, w, h: size > 12 ? 0.328 : 0.303, fontSize: size, color: WHITE, align: align || 'left' });
  text(s, sub, { x, y: subY, w, h: size > 12 ? 0.328 : 0.303, fontSize: size, color: WHITE, align: align || 'left' });
}

// =============================================================== slides =====

// 1 - cover ("Discover." is white-on-white in the source deck).
function slide01(s) {
  title(s, [['Discover.']], 1.136, 3.0, 7.729, 1.5, 66, WHITE, 'center');
  slideChrome(s, { pillY: 0.194, logoY: 0.149, page: '1', pageY: 7.012, pageW: 0.502, socialY: 7.053 });
}

// 2 - section break: welcome
function slide02(s) {
  rect(s, 0, 0, 10, 7.5, SCRIM);
  title(s, [['We Bring You To '], ['The New', BLUE], [' World.']], 1.453, 2.312, 3.172, 2.869, 40.5, WHITE);
  sectionLabels(s, 1.453, 2.518, 4.672, 3.172, 'Welcome To Our Company', 12);
  slideChrome(s, { pillY: 0.194, logoY: 0.149, page: '2', pageY: 7.012, pageW: 0.502, socialY: 7.053 });
}

// 3 - table of contents
function slide03(s) {
  rect(s, 0, -0.006, 4.297, 7.511, SCRIM);
  title(s, [['Table Of '], ['Content.', BLUE]], 0.602, 2.792, 3.094, 1.911, 40.5, WHITE);
  const items = ['Welcome Message', 'About Us', 'Meet The Team', 'Our Works', 'Process Data'];
  items.forEach(function (name, i) {
    const y = 1.712 + i * 0.833;
    text(s, String(i + 1).padStart(2, '0'), {
      x: 4.898, y: y + 0.046, w: 0.797, h: 0.656, fontSize: 33, fontFace: HEAD, color: BLUE, align: 'right',
    });
    label(s, name, 5.731, y, 3.414, 0.303);
    para(s, L_SHORT, 5.731, y + 0.216, 3.414, 0.529);
  });
  slideChrome(s, { logoY: 0.149, page: '3', pageY: 7.012, pageW: 0.502 });
}

// 4 - welcome message
function slide04(s) {
  rect(s, 0, -0.011, 10, 5.483, SCRIM);
  title(s, [['We Bring You '], ['To The ', BLUE], ['New World.']], 4.725, 1.512, 4.229, 1.534, 30, WHITE);
  para(s, L_LONG, 4.725, 2.784, 4.229, 0.983, { color: WHITE });
  rect(s, 5.0, 5.471, 5.0, 2.029, { color: BLUE });
  label(s, 'Welcome Message', 0.793, 6.091, 3.414, 0.303, { align: 'center' });
  para(s, L_SHORT, 0.793, 6.319, 3.414, 0.529, { align: 'center' });
  text(s, L_SHORT, { x: 5.492, y: 6.167, w: 4.015, h: 0.707, color: WHITE, italic: true, align: 'center' });
  slideChrome(s, { pillY: 0.2, logoY: 0.155 });
}

// 5 - section break: about
function slide05(s) {
  rect(s, 0, 0, 10, 7.5, SCRIM);
  title(s, [['About Our '], ['Company.', BLUE]], 2.05, 2.9, 5.9, 1.694, 40.5, WHITE, 'center');
  sectionLabels(s, 3.414, 3.1, 4.068, 3.172, 'Company Introduction', 13.5, 'center');
  slideChrome(s, { pillY: 0.2, logoY: 0.155, page: '5', pageY: 7.009, pageW: 0.502, socialY: 7.05 });
}

// 6 - capabilities
function slide06(s) {
  rect(s, 2.0, 0, 3.176, 7.5, { color: BLUE });
  title(s, [['Your '], ['Capabilities Are ', BLUE], ['Unlimited.']], 5.838, 1.778, 3.571, 1.404, 30, BLACK);
  para(s, L_MED, 5.838, 2.982, 3.571, 0.983);
  [['$1,356', 'Accommodation', 4.7], ['$2,783', 'Leisure', 7.134]].forEach(function (c) {
    card(s, c[2], 4.393, 2.275, 1.217);
    stat(s, c[0], c[2] + 0.232, 4.685, 1.812, { align: 'center' });
    label(s, c[1], c[2] + 0.232, 5.034, 1.812, 0.303, { align: 'center' });
  });
}

// 7 - "Bring The Best Things"
function slide07(s) {
  rect(s, 0, 0, 10, 5.035, SCRIM);
  title(s, [['Bring', BLUE], [' The Best Things '], ['You Ever', BLUE], [' Made.']], 1.055, 1.459, 2.978, 2.082, 30, WHITE);
  para(s, L_TELLUS, 4.359, 2.026, 4.586, 0.983, { color: WHITE });
  card(s, 0.719, 4.091, 4.179, 1.777);
  card(s, 5.102, 4.091, 4.179, 1.777);
  icon(s, 'car', 2.642, 4.414, 0.332, 0.25, BLUE);
  icon(s, 'food', 7.052, 4.34, 0.279, 0.335, BLUE);
  [['Transportation', 1.101], ['Food And Beverages', 5.485]].forEach(function (c) {
    label(s, c[0], c[1], 4.807, 3.414, 0.303, { align: 'center' });
    para(s, L_SHORT, c[1], 5.035, 3.414, 0.529, { align: 'center' });
  });
  slideChrome(s, { pillX: 8.908, pillY: 0.194, logoX: 0.181, logoY: 0.149 });
}

// 8 - creativity + stat grid
function slide08(s) {
  rect(s, 5.0, 0, 5.0, 7.5, SCRIM);
  title(s, [['Creativity ', BLUE], ['Is The Basic Rule Of '], ['Everything.', BLUE]], 5.781, 2.709, 3.438, 2.082, 30, WHITE);
  para(s, L_MED, 0.781, 2.082, 3.571, 0.983);
  const grid = [['$1,356', 'Accommodation', 0.781, 3.194], ['$785', 'Leisure', 2.603, 3.194],
    ['154K', 'Visitor', 0.781, 3.972], ['363K', 'Resident', 2.603, 3.972]];
  grid.forEach(function (g) {
    stat(s, g[0], g[2], g[3], 1.75);
    label(s, g[1], g[2], g[3] + 0.349, 1.75, 0.303);
  });
  divider(s, 2.46, 3.269, 0.51);
  divider(s, 2.46, 4.078, 0.51);
  label(s, 'Food And Beverages', 0.781, 4.75, 3.571, 0.303);
  para(s, L_SHORT, 0.781, 4.978, 3.571, 0.529);
  slideChrome(s, { pillY: 0.155, socialY: 7.087 });
}

// 9 - checklist card
function slide09(s) {
  rect(s, 5.213, 0, 3.176, 7.5, { color: BLUE });
  title(s, [['We '], ['Make The ', BLUE], ['Best Things '], ['We Can ', BLUE], ['Do']], 0.75, 1.984, 2.952, 1.603, 30, BLACK);
  para(s, L_MED, 0.75, 3.515, 2.952, 1.21, { align: 'justify' });
  pill(s, 0.75, 5.008, 0.871, 0.312, { color: BLUE });
  text(s, 'Travel', { x: 0.904, y: 5.038, w: 0.563, h: 0.252, fontSize: 9, fontFace: HEAD, color: WHITE, align: 'center', wrap: false });
  pill(s, 1.744, 5.003, 0.871, 0.312, { color: WHITE }, { color: BLUE, width: 1 });
  text(s, 'Next', { x: 1.945, y: 5.033, w: 0.468, h: 0.252, fontSize: 9, fontFace: HEAD, color: BLUE, align: 'center', wrap: false });
  card(s, 6.207, 3.354, 3.305, 2.435, 0.158);
  ['Camp Spot', 'Leisure Activities', 'Beautiful Scenery', 'Good Temperature'].forEach(function (t, i) {
    icon(s, 'check', 6.796, 3.735 + i * 0.4663, 0.269, 0.269, BLUE);
    label(s, t, 7.181, 3.719 + i * 0.466, 1.947, 0.303);
  });
}

// 10 - three service cards
function slide10(s) {
  rect(s, 0, 0, 2.75, 7.5, SCRIM);
  title(s, [['Do Your Best '], ['At Your', BLUE], [' Work']], 5.636, 2.032, 3.703, 1.183, 30, BLACK);
  const cards = [['Transportation', 'car', 1.349, 2.102, 0.332, 0.25], ['Accommodation', 'food', 1.375, 3.583, 0.279, 0.335],
    ['Leisure', 'beach', 1.375, 5.125, 0.279, 0.296]];
  cards.forEach(function (c, i) {
    const y = 1.54 + i * 1.5225;
    card(s, 0.906, y, 4.072, 1.375);
    icon(s, c[1], c[2], c[3], c[4], c[5], BLUE);
    label(s, c[0], 2.014, y + 0.203, 2.658, 0.303);
    para(s, L_CONGUE, 2.014, y + 0.416, 2.658, 0.756);
  });
  para(s, L_VIVERRA, 5.636, 3.133, 3.703, 0.983);
  text(s, L_SHORT, { x: 5.636, y: 4.242, w: 3.703, h: 0.707, italic: true });
  pill(s, 5.636, 5.194, 0.871, 0.312, { color: BLUE });
  text(s, 'Travel', { x: 5.79, y: 5.224, w: 0.563, h: 0.252, fontSize: 9, fontFace: HEAD, color: WHITE, align: 'center', wrap: false });
  pill(s, 6.63, 5.189, 0.871, 0.312, { color: WHITE }, { color: BLUE, width: 1 });
  text(s, 'Next', { x: 6.831, y: 5.219, w: 0.468, h: 0.252, fontSize: 9, fontFace: HEAD, color: BLUE, align: 'center', wrap: false });
  slideChrome(s, { logoY: 0.146, page: '10', pageY: 7.01 });
}

// -- timeline helpers shared by slides 11 and 12 -----------------------------
function timelineNode(s, y) { // small ring on the spine
  s.addShape('ellipse', { x: 4.94, y, w: 0.121, h: 0.121, fill: { color: WHITE }, line: { color: BLUE, width: 1.5 } });
}
function timelineSpine(s, y, h) {
  s.addShape('line', { x: 5.0, y, w: 0, h, line: { color: BLUE, width: 1.5 } });
}
function timelineYear(s, year, y) {
  text(s, year, { x: 4.516, y, w: 0.969, h: 0.353, fontSize: 15, fontFace: HEAD, color: BLUE, align: 'center' });
}
function timelineEntry(s, heading, x, y, align) {
  label(s, heading, x, y, 3.26, 0.303, { align });
  para(s, L_PORTTITOR, x, y + 0.213, 3.26, 0.756, { align });
}

// 11 - company history (top half)
function slide11(s) {
  title(s, [['Company '], ['History.', BLUE]], 1.344, 0.787, 7.311, 0.915, 30, BLACK, 'center');
  s.addShape('ellipse', { x: 4.826, y: 2.378, w: 0.348, h: 0.348, fill: { color: WHITE }, line: { color: BLUE, width: 1.5 } });
  timelineSpine(s, 2.726, 0.585);
  [3.311, 3.887, 4.756, 5.332].forEach(function (y) { timelineNode(s, y); });
  timelineSpine(s, 4.008, 0.749);
  timelineSpine(s, 5.453, 2.047);
  timelineYear(s, '2015', 3.483);
  timelineYear(s, '2016', 4.928);
  timelineEntry(s, 'Beginning', 5.838, 3.175, 'left');
  timelineEntry(s, 'Investment', 0.901, 4.62, 'right');
  iconTile(s, 3.619, 3.388, 'car', 3.76, 3.561, 0.26, 0.196);
  iconTile(s, 5.838, 4.833, 'laptop', 5.987, 5.021, 0.244, 0.168);
}

// 12 - company history (bottom half)
function slide12(s) {
  timelineSpine(s, 0, 2.047);
  timelineSpine(s, 2.744, 0.658);
  timelineSpine(s, 4.098, 0.658);
  timelineSpine(s, 5.453, 0.373);
  [2.047, 2.623, 3.402, 3.978, 4.756, 5.332].forEach(function (y) { timelineNode(s, y); });
  s.addShape('ellipse', { x: 4.826, y: 5.825, w: 0.348, h: 0.348, fill: { color: WHITE }, line: { color: BLUE, width: 1.5 } });
  timelineYear(s, '2017', 2.219);
  timelineYear(s, '2018', 3.573);
  timelineYear(s, '2019', 4.928);
  timelineEntry(s, 'Employee', 5.838, 1.911, 'left');
  timelineEntry(s, 'Branch', 0.901, 3.265, 'right');
  timelineEntry(s, 'Success', 5.838, 4.62, 'left');
  iconTile(s, 3.619, 2.124, 'bed', 3.76, 2.294, 0.26, 0.208);
  iconTile(s, 5.838, 3.479, 'beach', 5.975, 3.608, 0.268, 0.284);
  iconTile(s, 3.619, 4.833, 'food', 3.76, 4.948, 0.26, 0.312);
}

// 13 - section break: team
function slide13(s) {
  rect(s, 3.672, 0, 6.328, 7.5, SCRIM);
  title(s, [['Meet '], ['Our ', BLUE], ['Precious '], ['Team.', BLUE]], 5.683, 2.312, 3.172, 2.869, 40.5, WHITE);
  sectionLabels(s, 5.683, 2.518, 4.672, 3.172, 'Meet The Team', 12);
  slideChrome(s, { pillY: 0.192, logoY: 0.147, page: '13', pageY: 7.005, socialY: 7.046 });
}

// 14 - team member profile
function slide14(s) {
  rect(s, 5.522, 0, 2.658, 7.5, { color: BLUE });
  title(s, [['Jennifer '], ['Bella.', BLUE]], 0.75, 1.833, 4.687, 0.947, 30, BLACK);
  label(s, 'Company Founder & Manager', 0.75, 2.525, 4.687, 0.303);
  para(s, L_VIVERRA, 0.75, 2.904, 4.687, 0.756);
  progress(s, 'Management', '80%', 0.75, 4.094, 3.812, 3.062, 1.938, 3.613);
  progress(s, 'Leadership', '50%', 0.75, 4.668, 3.812, 2.013, 1.938, 3.613);
  progress(s, 'Creativity', '90%', 0.75, 5.242, 3.812, 3.381, 1.938, 3.613);
  card(s, 4.967, 4.02, 4.072, 1.375);
  icon(s, 'car', 5.411, 4.582, 0.332, 0.25, BLUE);
  label(s, 'Ideas Program', 6.075, 4.223, 2.658, 0.303);
  para(s, L_CONGUE, 6.075, 4.436, 2.658, 0.756);
}

// 15 - two-up team cards
function slide15(s) {
  rect(s, 0, 0, 4.272, 7.5, { color: BLUE });
  title(s, [['Meet ', BLUE], ['Our Best And '], ['Amazing Team.', BLUE]], 6.234, 2.312, 3.04, 1.657, 30, BLACK);
  para(s, L_VIVERRA, 6.234, 3.916, 3.04, 1.21);
  [['Maria Jessie', 1.009, 4.984, 1.145, 5.131, 1.704], ['Claire Bella', 3.553, 4.979, 3.689, 5.126, 4.25]].forEach(function (m) {
    card(s, m[1], m[2], 1.98, 0.942);
    label(s, m[0], m[3], m[4], 1.707, 0.303, { align: 'center' });
    text(s, 'General Manager', { x: m[3], y: m[4] + 0.217, w: 1.707, h: 0.252, fontSize: 9, align: 'center' });
    socialRow(s, m[5], 5.608, 0.145, BLUE);
  });
  slideChrome(s, { logoY: 0.181, page: '13', pageY: 7.01, pageW: 0.631 });
}

// 16 - three team photos (photo wells are empty in the source deck)
function slide16(s) {
  title(s, [['We Are ', BLUE], ['Great And '], ['Amazing Team.', BLUE]], 6.265, 1.354, 3.04, 1.616, 30, BLACK);
  para(s, L_VIVERRA, 0.743, 1.458, 4.687, 0.756, { align: 'right' });
  text(s, L_SHORT, { x: 0.743, y: 2.318, w: 4.687, h: 0.505, italic: true, align: 'right' });
  [['Maria Jessie', 0.284, 0.369], ['Claire Bella', 3.616, 3.7], ['Jenifer Clint', 6.947, 7.031]].forEach(function (m) {
    label(s, m[0], m[1], 5.644, 1.707, 0.303, { color: WHITE });
    text(s, 'General Manager', { x: m[1], y: 5.86, w: 1.707, h: 0.252, fontSize: 9, color: WHITE });
    socialRow(s, m[2], 6.12, 0.145, WHITE);
  });
}

// 17 - section break: works
function slide17(s) {
  rect(s, 0, -0.006, 6.891, 7.507, SCRIM);
  title(s, [['Here', BLUE], [' Some Of Our '], ['Works.', BLUE]], 1.375, 3.747, 4.745, 1.697, 40.5, WHITE);
  sectionLabels(s, 1.375, 3.659, 5.186, 3.172, 'Company Portfolio', 12);
  slideChrome(s, { pillY: 0.185, logoY: 0.14, page: '17', pageY: 7.012, socialY: 7.053 });
}

// 18 - single portfolio picture
function slide18(s) {
  rect(s, 4.624, 0, 1.978, 7.5, { color: BLUE });
  title(s, [['Our '], ['Amazing', BLUE], [' Portfolio '], ['Picture.', BLUE]], 0.726, 1.252, 2.711, 2.194, 30, BLACK);
  card(s, 0.726, 3.461, 4.586, 2.162);
  iconTile(s, 1.058, 3.755, 'car', 1.2, 3.928, 0.26, 0.196);
  label(s, 'Amazing Portfolio Picture', 1.058, 4.325, 3.922, 0.303);
  para(s, L_LIBERO, 1.058, 4.529, 3.922, 0.756, { align: 'justify' });
}

// 19 - two picture portfolio
function slide19(s) {
  rect(s, 5.17, 4.882, 4.83, 2.618, { color: BLUE });
  title(s, [['Two '], ['Picture ', BLUE], ['Portfolio '], ['Slide.', BLUE]], 6.336, 1.223, 2.711, 2.194, 30, BLACK);
  para(s, L_MALESUADA, 0.81, 3.911, 2.786, 0.983, { align: 'justify' });
  stat(s, '$1,356', 0.81, 5.023, 1.196);
  label(s, 'Hotel', 0.81, 5.372, 1.196, 0.303);
  stat(s, '$785', 2.225, 5.023, 1.196);
  label(s, 'Leisure', 2.225, 5.372, 1.196, 0.303);
  divider(s, 2.082, 5.098, 0.51);
  slideChrome(s, { socialY: 7.098 });
}

// 20 - three picture portfolio
function slide20(s) {
  rect(s, 7.1, 2.559, 2.9, 4.941, { color: BLUE });
  title(s, [['Three '], ['Great', BLUE], [' Picture'], [' Portfolio.', BLUE]], 3.655, 4.338, 2.69, 2.194, 30, BLACK);
  para(s, L_EROS, 0.953, 4.651, 2.583, 1.438, { align: 'right' });
  iconTile(s, 3.619, 2.386, 'car', 3.76, 2.559, 0.26, 0.196);
  slideChrome(s, { socialY: 7.101 });
}

// 21 - phone mockups
function slide21(s) {
  rect(s, 0, 0, 2.291, 4.328, { color: BLUE });
  title(s, [['Device '], ['Mockup Slide.', BLUE]], 6.188, 1.905, 3.126, 1.165, 30, BLACK);
  para(s, L_AMET, 6.188, 3.009, 3.126, 0.983, { align: 'justify' });
  const grid = [['$1,356', 'Accommodation', 6.189, 4.119], ['$785', 'Leisure', 8.01, 4.119],
    ['154K', 'Visitor', 6.189, 4.897], ['363K', 'Resident', 8.01, 4.897]];
  grid.forEach(function (g) {
    stat(s, g[0], g[2], g[3], 1.75);
    label(s, g[1], g[2], g[3] + 0.349, 1.75, 0.303);
  });
  divider(s, 7.867, 4.194, 0.51);
  divider(s, 7.867, 5.003, 0.51);
  phoneMockup(s, 0.891, 0.545, 2.095, 4.504);
  phoneMockup(s, 3.25, 2.597, 2.095, 4.504);
  label(s, 'Device One', 0.891, 5.253, 2.095, 0.303);
  para(s, L_ELIT, 0.891, 5.449, 2.095, 0.529, { align: 'justify' });
  label(s, 'Device Two', 3.25, 1.636, 2.095, 0.303);
  para(s, L_ELIT, 3.25, 1.832, 2.095, 0.529, { align: 'justify' });
}

// 22 - laptop mockup
function slide22(s) {
  rect(s, 0, 0, 3.528, 7.5, { color: BLUE });
  laptopMockup(s, 0.345, 2.359, 4.605, 2.781);
  title(s, [['Amazing '], ['Device Mockup', BLUE], [' Slide.']], 5.328, 1.85, 4.047, 1.157, 30, BLACK);
  para(s, L_AMET, 5.328, 2.89, 4.047, 0.756, { align: 'justify' });
  label(s, 'Amazing Portfolio Picture', 5.328, 3.716, 4.047, 0.303);
  para(s, L_PULVINAR, 5.328, 3.924, 4.047, 0.529, { align: 'justify' });
  progress(s, 'Management', '80%', 5.328, 4.926, 3.812, 3.062, 1.938, 8.191);
  progress(s, 'Leadership', '50%', 5.328, 5.5, 3.812, 2.013, 1.938, 8.191);
  slideChrome(s, { logoY: 0.189, page: '22', pageY: 7.008 });
}

// 23 - testimonials
function slide23(s) {
  rect(s, 0, 0, 4.703, 7.496, SCRIM);
  title(s, [['Some ', BLUE], ['Testimonial From '], ['User.', BLUE]], 0.999, 2.653, 2.69, 2.194, 30, WHITE);
  const quotes = [['Jessica', 5.297, 1.663, 6.579, 1.809, 6.655, 2.558],
    ['Amanda', 4.422, 3.082, 5.702, 3.235, 5.775, 3.974],
    ['Hilton', 5.297, 4.505, 6.579, 4.661, 6.655, 5.389]];
  quotes.forEach(function (q) {
    card(s, q[1], q[2], 4.141, 1.332, 0.139);
    label(s, q[0], q[3], q[4], 2.642, 0.303);
    para(s, L_TINY, q[3], q[4] + 0.209, 2.642, 0.529);
    stars(s, q[5], q[6]);
  });
  slideChrome(s, { logoY: 0.174, page: '23', pageY: 7.024 });
}

// 24 - section break: data
function slide24(s) {
  rect(s, 0, 0, 10, 4.781, SCRIM);
  title(s, [['Process '], ['Data.', BLUE]], 2.533, 2.175, 4.934, 1.17, 40.5, WHITE, 'center');
  sectionLabels(s, 3.414, 2.173, 2.957, 3.172, 'Chart Data & Infographics', 12, 'center');
  slideChrome(s, { pillY: 0.188, logoY: 0.143, page: '24', pageY: 7.013, socialY: 7.054 });
}

// 25 - stacked-slab infographic
// Each slab is three faces of a 3D box: dark top diamond, base-colour left, light right.
const SLAB_TOP = [[0.497, 0], [1, 0.275], [0.497, 0.540], [0, 0.275]];
const SLAB_LEFT = [[0, 0.275], [0.497, 0.540], [0.497, 1], [0, 0.735]];
const SLAB_RIGHT = [[1, 0.275], [0.497, 0.540], [0.497, 1], [1, 0.735]];
const SLAB_COLORS = [ // [base, dark top (lumMod 75%), light right (lumMod 60 / lumOff 40)]
  [BLUE, '004FC3', '69A6FF'], [BLUE2, '026BC1', '6ABBFE'],
  [BLUE3, '0079C3', '69C6FF'], [BLUE4, '098CBA', '70D4F8'],
];

function slab(s, x, y, w, h, colors) {
  [[SLAB_TOP, colors[1]], [SLAB_LEFT, colors[0]], [SLAB_RIGHT, colors[2]]].forEach(function (face) {
    s.addShape('custGeom', {
      x, y, w, h, fill: { color: face[1] }, line: { type: 'none' },
      points: face[0].map(function (p, i) {
        return { x: p[0] * w, y: p[1] * h, moveTo: i === 0 };
      }).concat([{ close: true }]),
    });
  });
}

function slide25(s) {
  title(s, [['Infographic '], ['Vector.', BLUE]], 1.344, 0.787, 7.311, 0.915, 30, BLACK, 'center');
  [[3.249, 3.106], [3.249, 4.53], [6.151, 3.841], [6.151, 5.261]].forEach(function (l) {
    s.addShape('line', { x: l[0], y: l[1], w: 0.6, h: 0, line: { color: GRAY, width: 2 } });
  });
  [2.606, 3.313, 4.01, 4.724].forEach(function (y, i) {
    slab(s, 3.751, y, 2.499, i === 3 ? 1.037 : 1.025, SLAB_COLORS[i]);
    text(s, 'Step ' + (i + 1), {
      x: 4.026, y: 3.102 + i * 0.708, w: 0.691, h: 0.303,
      fontSize: 12, fontFace: HEAD, color: WHITE, rotate: 11.71, valign: 'top', wrap: false,
    });
  });
  const steps = [ // [ring x, ring y, ring colour, icon, icon x, icon y, icon w, icon h, label, label x, label y, align]
    [2.583, 2.746, BLUE, 'bed', 2.815, 3.019, 0.26, 0.208, 'Step One', 0.402, 2.593, 'right'],
    [6.693, 3.477, BLUE2, 'beach', 6.914, 3.697, 0.268, 0.284, 'Step Two', 7.714, 3.335, 'left'],
    [2.583, 4.163, BLUE3, 'car', 2.815, 4.432, 0.26, 0.196, 'Step Three', 0.402, 4.01, 'right'],
    [6.693, 4.897, BLUE4, 'food', 6.924, 5.103, 0.26, 0.312, 'Step Four', 7.714, 4.751, 'left'],
  ];
  steps.forEach(function (st) {
    s.addShape('ellipse', { x: st[0], y: st[1], w: 0.724, h: 0.724, fill: { color: WHITE }, line: { color: st[2], width: 2 } });
    icon(s, st[3], st[4], st[5], st[6], st[7], st[2]);
    label(s, st[8], st[9], st[10], 1.884, 0.303, { align: st[11] });
    para(s, L_TINY, st[9], st[10] + 0.226, 1.884, 0.756, { align: st[11] });
  });
}

// 26 - snaking arrow roadmap
// U-turn arrow drawn as a custom path: two concentric half-ellipses joined by a
// straight leg on the left and a stem + arrowhead on the right. Fractions of w/h.
const U_LEG_OUT = 0.003, U_LEG_IN = 0.187, U_STEM_IN = 0.728, U_STEM_OUT = 0.908;
const U_ARC_Y = 0.46, U_INNER_TOP = 0.199, U_HEAD_Y = 0.757;
const K = 0.5523; // circle-to-Bezier constant

function uturn(s, x, y, w, h, color, flipped) {
  const oCx = (U_LEG_OUT + U_STEM_OUT) / 2, oRx = (U_STEM_OUT - U_LEG_OUT) / 2;
  const iCx = (U_LEG_IN + U_STEM_IN) / 2, iRx = (U_STEM_IN - U_LEG_IN) / 2;
  const iRy = U_ARC_Y - U_INNER_TOP;
  const P = function (fx, fy) { return { x: fx * w, y: fy * h }; };
  const cubic = function (fx, fy, x1, y1, x2, y2) {
    return { x: fx * w, y: fy * h, curve: { type: 'cubic', x1: x1 * w, y1: y1 * h, x2: x2 * w, y2: y2 * h } };
  };
  const opt = {
    x, y, w, h, fill: { color }, line: { type: 'none' },
    points: [
      Object.assign(P(U_LEG_OUT, 1), { moveTo: true }),
      P(U_LEG_OUT, U_ARC_Y),
      cubic(oCx, 0, U_LEG_OUT, U_ARC_Y * (1 - K), oCx - oRx * K, 0),
      cubic(U_STEM_OUT, U_ARC_Y, oCx + oRx * K, 0, U_STEM_OUT, U_ARC_Y * (1 - K)),
      P(U_STEM_OUT, U_HEAD_Y), P(1, U_HEAD_Y), P(0.818, 0.99), P(0.634, U_HEAD_Y),
      P(U_STEM_IN, U_HEAD_Y), P(U_STEM_IN, U_ARC_Y),
      cubic(iCx, U_INNER_TOP, U_STEM_IN, U_ARC_Y - iRy * K, iCx + iRx * K, U_INNER_TOP),
      cubic(U_LEG_IN, U_ARC_Y, iCx - iRx * K, U_INNER_TOP, U_LEG_IN, U_ARC_Y - iRy * K),
      P(U_LEG_IN, 1),
      { close: true },
    ],
  };
  if (flipped) { opt.rotate = 180; opt.flipH = true; }
  s.addShape('custGeom', opt);
}

function slide26(s) {
  title(s, [['Infographic '], ['Vector.', BLUE]], 1.344, 0.787, 7.311, 0.915, 30, BLACK, 'center');
  // [arrow x, arrow y, colour, flipped?, icon, icon x, icon y, icon w, icon h]
  const arrows = [
    [1.232, 2.544, BLUE, false, 'laptop', 1.998, 3.362, 0.244, 0.168],
    [2.627, 3.938, BLUE2, true, 'car', 3.387, 4.757, 0.26, 0.196],
    [4.021, 2.544, BLUE3, false, 'bed', 4.777, 3.33, 0.26, 0.208],
    [5.416, 3.938, BLUE4, true, 'food', 6.176, 4.697, 0.26, 0.312],
    [6.81, 2.544, BLUE5, false, 'beach', 7.566, 3.304, 0.268, 0.284],
  ];
  arrows.forEach(function (a) {
    uturn(s, a[0], a[1], 1.958, 1.807, a[2], a[3]);
    s.addShape('ellipse', {
      x: a[0] + 0.485, y: a[1] + (a[3] ? 0.512 : 0.485), w: 0.81, h: 0.81,
      fill: { color: a[2] }, line: { color: WHITE, width: 1.5 },
    });
    icon(s, a[4], a[5], a[6], a[7], a[8], WHITE, a[2]);
    // Straight stand-in for the reference's arch-warped caption.
    text(s, 'Add title here', {
      x: a[0] - 0.388, y: a[1] + (a[3] ? 0.934 : 0.573), w: 1.52, h: 0.3,
      fontSize: 12, fontFace: HEAD, color: WHITE, rotate: a[3] ? 52 : -52, align: 'center', wrap: false,
    });
  });
  [['2012', 1.823, BLUE], ['2014', 3.22, BLUE2], ['2016', 4.61, BLUE3], ['2018', 6.009, BLUE4], ['2020', 7.403, BLUE5]]
    .forEach(function (y) {
      text(s, y[0], { x: y[1], y: 3.98, w: 0.595, h: 0.303, fontSize: 12, fontFace: HEAD, color: y[2], wrap: false });
    });
  [['Start Project', 0.562, 1.719, 0.566], ['End Project', 7.623, 1.612, 7.574]].forEach(function (e) {
    text(s, e[0], { x: e[1], y: 4.55, w: e[2], h: 0.303, align: 'center', valign: 'top' });
    para(s, 'Sed perspiciatis unde omnis voluptatem fringilla.', e[3], 4.774, 1.71, 0.756, { align: 'center', valign: 'top' });
  });
}

// 27 - stacked bar chart
function slide27(s) {
  title(s, [['Chart '], ['Data.', BLUE]], 1.344, 0.787, 7.311, 0.915, 30, BLACK, 'center');
  const cats = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  s.addChart('bar', [
    { name: 'Series 1', labels: cats, values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', labels: cats, values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', labels: cats, values: [2, 2, 3, 5] },
  ], {
    x: 0.604, y: 2.214, w: 5.099, h: 2.265,
    barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 150, barOverlapPct: 100,
    chartColors: [BLUE, BLUE2, BLUE3],
    showLegend: false, showTitle: false, showValue: false,
    catAxisLineColor: GRAY, catAxisLabelColor: '595959', catAxisLabelFontFace: BODY, catAxisLabelFontSize: 12,
    catAxisMajorTickMark: 'none', catAxisMinorTickMark: 'none',
    valAxisLineShow: false, valAxisLabelColor: '595959', valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12,
    valAxisMajorTickMark: 'none', valAxisMinorTickMark: 'none', valGridLine: { style: 'none' },
  });
  para(s, L_AMET, 5.943, 2.361, 3.453, 0.983, { align: 'justify' });
  label(s, 'Chart Data', 5.943, 3.51, 3.453, 0.303);
  para(s, L_PULVINAR, 5.943, 3.72, 3.453, 0.756, { align: 'justify' });
  const legend = [['Series One', BLUE, 0.818, 'car', 0.959, 5.271, 0.26, 0.196, 1.458],
    ['Series Two', BLUE2, 3.738, 'bed', 3.879, 5.271, 0.26, 0.208, 4.377],
    ['Series Three', BLUE3, 6.655, 'beach', 6.789, 5.195, 0.268, 0.284, 7.296]];
  legend.forEach(function (g) {
    round(s, g[2], 5.098, 0.543, 0.543, { color: g[1] }, 0.18);
    icon(s, g[3], g[4], g[5], g[6], g[7], WHITE, g[1]);
    label(s, g[0], g[8], 4.875, 1.885, 0.303);
    para(s, L_TINY, g[8], 5.086, 1.885, 0.756, { align: 'justify' });
  });
}

// 28 - US map with percentage callouts
// Mainland silhouette traced from the reference art, in 0..1 box coordinates.
const US_TOP = [[0.03, 0.28], [0.06, 0.30], [0.09, 0.02], [0.12, 0.00], [0.15, 0.02], [0.18, 0.03], [0.21, 0.04],
  [0.23, 0.05], [0.26, 0.06], [0.29, 0.06], [0.32, 0.07], [0.35, 0.08], [0.38, 0.08], [0.41, 0.09], [0.44, 0.09],
  [0.47, 0.10], [0.50, 0.10], [0.53, 0.22], [0.56, 0.22], [0.59, 0.22], [0.62, 0.25], [0.65, 0.22], [0.68, 0.18],
  [0.70, 0.17], [0.73, 0.34], [0.76, 0.35], [0.79, 0.32], [0.82, 0.26], [0.85, 0.25], [0.88, 0.18], [0.91, 0.17],
  [0.94, 0.14], [0.97, 0.07], [1.00, 0.14]];
const US_BOTTOM = [[1.00, 0.15], [0.97, 0.27], [0.94, 0.29], [0.91, 0.36], [0.88, 0.41], [0.85, 0.91], [0.82, 0.89],
  [0.79, 0.79], [0.76, 0.75], [0.73, 0.76], [0.70, 0.75], [0.68, 0.75], [0.65, 0.79], [0.62, 0.79], [0.59, 0.79],
  [0.56, 0.79], [0.53, 0.82], [0.50, 0.84], [0.47, 0.92], [0.44, 0.86], [0.41, 0.79], [0.38, 0.79], [0.35, 0.79],
  [0.32, 0.71], [0.29, 0.69], [0.26, 0.70], [0.23, 0.69], [0.21, 0.67], [0.18, 0.64], [0.15, 0.62], [0.12, 0.61],
  [0.09, 0.55], [0.06, 0.52], [0.03, 0.29]];
const ALASKA = [[0.05, 0.78], [0.13, 0.72], [0.17, 0.80], [0.22, 0.78], [0.20, 0.90], [0.10, 1.00], [0.00, 0.95]];

function polygon(s, pts, x, y, w, h, color) {
  s.addShape('custGeom', {
    x, y, w, h, fill: { color }, line: { type: 'none' },
    points: pts.map(function (p, i) { return { x: p[0] * w, y: p[1] * h, moveTo: i === 0 }; }).concat([{ close: true }]),
  });
}

function slide28(s) {
  title(s, [['United '], ['States.', BLUE]], 1.344, 0.787, 7.311, 0.915, 30, BLACK, 'center');
  polygon(s, US_TOP.concat(US_BOTTOM), 0.775, 2.629, 4.615, 3.017, MAPGRAY);
  polygon(s, ALASKA, 0.775, 2.629, 4.615, 3.017, MAPGRAY);
  [['20%', 0.912, 3.014, BLUE], ['40%', 3.102, 2.773, BLUE2], ['60%', 2.255, 4.097, BLUE3], ['90%', 4.242, 3.993, BLUE4]]
    .forEach(function (c) {
      s.addShape('wedgeRoundRectCallout', { x: c[1], y: c[2], w: 0.75, h: 0.502, fill: { color: c[3] }, line: { type: 'none' } });
      text(s, c[0], { x: c[1], y: c[2] + 0.035, w: 0.75, h: 0.438, fontSize: 20, fontFace: HEAD, color: WHITE, align: 'center' });
    });
  stat(s, '12,535,248', 5.772, 2.616, 3.453);
  text(s, 'Total States Population That Using Fast Speed Internet', { x: 5.772, y: 3.047, w: 3.453, h: 0.505 });
  para(s, L_AMET, 5.772, 3.595, 3.453, 0.983, { align: 'justify' });
  label(s, 'United States', 5.772, 4.693, 3.453, 0.303);
  para(s, L_PULVINAR, 5.772, 4.903, 3.453, 0.756, { align: 'justify' });
}

// 29 - contact
function slide29(s) {
  title(s, [['Contact Us']], 6.143, 1.511, 3.274, 0.915, 30, BLACK);
  text(s, 'Address', { x: 6.143, y: 2.768, w: 3.274, h: 0.303, valign: 'top' });
  para(s, '153 Central Blvd #34-54 Telok Ayer , Downtown\nSingapore 125325', 6.143, 2.992, 3.274, 0.529, { valign: 'top' });
  text(s, 'Phone Number', { x: 6.143, y: 3.611, w: 3.274, h: 0.303, valign: 'top' });
  para(s, '+1 234 567 890\n+1 098 765 432', 6.143, 3.835, 3.274, 0.529, { valign: 'top' });
  [['f', 4.481, 'Facebook.com/Username'], ['in', 4.886, 'Linked.in/12301287/Username'], ['t', 5.288, 'Twitter.com/Username']]
    .forEach(function (r) {
      social(s, 6.19, r[1], 0.217, r[0], BLUE);
      text(s, r[2], { x: 6.555, y: r[1] - 0.006, w: 2.862, h: 0.252, fontSize: 9, valign: 'top' });
    });
  pill(s, 2.667, 2.603, 0.985, 0.309, { color: BLUE });
  text(s, 'Location', { x: 2.702, y: 2.618, w: 0.915, h: 0.303, color: WHITE, align: 'center', valign: 'top' });
  icon(s, 'pin', 3.022, 3.004, 0.275, 0.332, BLUE);
}

// 30 - thank you
function slide30(s) {
  rect(s, 0, 1.049, 10, 5.402, SCRIM);
  title(s, [['Thank '], ['You.', BLUE]], 1.136, 3.0, 7.729, 1.5, 66, WHITE, 'center');
  text(s, 'See You Soon', { x: 3.414, y: 4.165, w: 3.172, h: 0.303, color: WHITE, align: 'center' });
}

// =============================================================== assembly ===
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
  slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

function build() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_4x3'; // 10 x 7.5 in
  pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
  pptx.title = 'Holiday Co. - Travel';

  BUILDERS.forEach(function (buildSlide, i) {
    const slide = pptx.addSlide();
    slide.background = { color: WHITE };
    masterChrome(slide, i + 1);
    buildSlide(slide);
  });

  return pptx.writeFile({ fileName: path.join(__dirname, '0bbba078-e6be-449e-b188-b918ff1eb452_grok_final.pptx') });
}

build().then(function (f) { console.log('wrote ' + f); }).catch(function (e) { console.error(e); process.exit(1); });
