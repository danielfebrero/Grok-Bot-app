/**
 * Art Gallery — 25-slide deck rebuilt with pptxgenjs.
 * Raster images in the source deck are represented by native-shape placeholders.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const GREEN     = '879C7D';   // theme accent5
const GREEN_D   = '64785B';   // accent5, 75% luminance
const GREEN_L   = 'E7EBE5';   // accent5, 20% luminance / 80% offset
const PURPLE    = '5E4D8D';   // theme accent6
const PURPLE_D  = '463A6A';   // accent6, 75% luminance
const PURPLE_XD = '2F2746';   // accent6, 50% luminance
const PURPLE_L  = '9B8EC1';   // accent6, 60%/40%
const PURPLE_XL = 'BDB3D6';
const WHITE     = 'FFFFFF';
const INK       = '404040';   // body copy on light backgrounds
const GREY_F2   = 'F2F2F2';
const GREY_D9   = 'D9D9D9';
const GREY_80   = '808080';

const HEAD = 'Lora';          // theme major font
const BODY = 'Raleway';       // theme minor font

const TITLE_SZ = 44;          // shared title size across the deck
const LEAD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet ' +
             'orci et consequat. Morbi semper eros vitae tincidunt porta. Mauris euismod. ' +
             'Cum sociis natoque penatibus.';
const LEAD2 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula';
const LEAD3 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula ' +
              'eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient ' +
              'montes, nascetur ridiculus mus. ';
const LEAD4 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ';

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: 40 / 3, height: 7.5 });
pptx.layout = 'DECK';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };

/* ------------------------------------------------------------- tiny helpers */
const S = (t) => pptx.ShapeType[t];

/** Solid-filled shape. */
function shape(slide, type, o) {
  slide.addShape(S(type), o);
}
/** Text box; every deck text box is top-anchored with the PowerPoint default insets. */
function text(slide, body, o) {
  slide.addText(body, Object.assign({ fontFace: BODY, fontSize: 18, color: '000000', valign: 'top' }, o));
}
/** Straight line between two points. */
function line(slide, x1, y1, x2, y2, o) {
  slide.addShape(S('line'), Object.assign({
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1,
  }, o));
}
/** Rectangle whose two bottom corners are rounded (the deck's green header panels). */
function panelRoundBottom(slide, x, y, w, h, r, o) {
  shape(slide, 'custGeom', Object.assign({
    x, y, w, h, points: [
      { x: 0, y: 0 }, { x: w, y: 0 }, { x: w, y: h - r },
      { x: w - r, y: h, curve: { type: 'cubic', x1: w, y1: h - r * 0.45, x2: w - r * 0.55, y2: h } },
      { x: r, y: h },
      { x: 0, y: h - r, curve: { type: 'cubic', x1: r * 0.55, y1: h, x2: 0, y2: h - r * 0.45 } },
      { close: true },
    ],
  }, o));
}
/** Polygon from a list of [x,y] pairs given in fractions of w/h. */
function poly(slide, x, y, w, h, pts, o) {
  shape(slide, 'custGeom', Object.assign({
    x, y, w, h,
    points: pts.map((p) => ({ x: p[0] * w, y: p[1] * h })).concat([{ close: true }]),
  }, o));
}

/* ------------------------------------------------------------- icon library */
/* Each icon is a list of primitives in a 0..1 box:
   r/e = outlined rect/ellipse, R/E = filled rect/ellipse, l = line, t = filled triangle. */
const ICONS = {
  museum:   [['t', .5, .02, .02, .28, .98, .28], ['R', .04, .30, .92, .07], ['R', .13, .40, .11, .40],
             ['R', .32, .40, .11, .40], ['R', .51, .40, .11, .40], ['R', .70, .40, .11, .40], ['R', .02, .85, .96, .10]],
  frame:    [['r', .03, .14, .94, .72], ['r', .12, .24, .76, .52], ['t', .22, .74, .44, .42, .66, .74],
             ['t', .52, .74, .70, .52, .88, .74], ['E', .66, .30, .11, .11]],
  easel:    [['e', .40, .02, .20, .14], ['r', .10, .12, .80, .50], ['l', .18, .70, .10, .98], ['l', .82, .70, .90, .98],
             ['l', .50, .70, .50, .98], ['l', .05, .66, .95, .66], ['t', .22, .56, .42, .30, .62, .56],
             ['t', .46, .56, .64, .38, .82, .56], ['E', .30, .22, .10, .10]],
  palette:  [['e', .02, .16, .78, .66], ['E', .16, .30, .11, .13], ['E', .34, .26, .11, .13],
             ['E', .52, .30, .11, .13], ['E', .30, .52, .11, .13], ['e', .56, .56, .40, .26],
             ['l', .60, .62, .92, .70], ['l', .60, .74, .92, .74]],
  masks:    [['e', .01, .12, .47, .76], ['e', .52, .12, .47, .76], ['E', .11, .34, .08, .09],
             ['E', .28, .34, .08, .09], ['E', .62, .34, .08, .09], ['E', .79, .34, .08, .09],
             ['l', .12, .62, .36, .70], ['l', .63, .70, .87, .62]],
  violin:   [['e', .04, .52, .42, .44], ['l', .34, .58, .96, .02], ['l', .44, .70, .99, .18],
             ['l', .16, .40, .44, .68], ['l', .40, .16, .68, .44]],
  flag:     [['l', .26, .02, .26, .98], ['r', .26, .08, .58, .40]],
  board:    [['r', .06, .06, .88, .60], ['l', .5, .66, .5, .96], ['l', .24, .96, .76, .96],
             ['l', .26, .30, .5, .22], ['l', .5, .22, .74, .30], ['l', .5, .22, .5, .52]],
  laptop:   [['r', .12, .12, .76, .50], ['R', .02, .68, .96, .09], ['l', .30, .32, .70, .32], ['l', .30, .45, .58, .45]],
  receipt:  [['r', .16, .03, .68, .84], ['l', .30, .26, .70, .26], ['l', .30, .42, .70, .42], ['l', .30, .58, .56, .58],
             ['t', .16, .87, .50, .99, .84, .87]],
  ticket:   [['r', .03, .26, .94, .48], ['l', .03, .44, .13, .50], ['l', .13, .50, .03, .56],
             ['l', .97, .44, .87, .50], ['l', .87, .50, .97, .56], ['l', .52, .30, .52, .70]],
  clipbrd:  [['r', .04, .09, .62, .88], ['R', .22, .02, .26, .12],
             ['r', .12, .26, .12, .12], ['r', .12, .46, .12, .12], ['r', .12, .66, .12, .12],
             ['l', .32, .30, .58, .30], ['l', .32, .50, .58, .50], ['l', .32, .70, .58, .70],
             ['r', .76, .22, .16, .58], ['t', .84, .12, .76, .26, .92, .26]],
  doc:      [['r', .06, .02, .62, .74], ['e', .18, .14, .22, .22], ['r', .28, .26, .22, .22],
             ['l', .16, .60, .50, .60], ['l', .16, .70, .50, .70],
             ['e', .52, .52, .46, .46], ['l', .75, .62, .75, .75], ['l', .75, .75, .86, .75]],
  head:     [['e', .14, .04, .60, .60], ['e', .32, .20, .28, .28], ['E', .41, .29, .10, .10],
             ['l', .16, .46, .16, .98], ['l', .16, .98, .58, .98], ['l', .58, .98, .58, .84],
             ['l', .58, .84, .82, .84], ['l', .82, .84, .82, .70], ['l', .82, .70, .70, .58]],
  checkC:   [['e', .02, .02, .92, .92], ['l', .22, .48, .42, .70], ['l', .42, .70, .96, .06]],
  mail:     [['r', .03, .20, .94, .60], ['l', .03, .22, .50, .56], ['l', .50, .56, .97, .22]],
  pin:      [['e', .18, .02, .64, .64], ['e', .36, .20, .28, .28], ['t', .50, .99, .26, .52, .74, .52]],
  phone:    [['E', .06, .06, .30, .30], ['E', .62, .62, .32, .32], ['l', .26, .30, .70, .70]],
  quote:    [['E', .02, .08, .40, .40], ['t', .06, .40, .42, .30, .26, .96],
             ['E', .56, .08, .40, .40], ['t', .60, .40, .96, .30, .80, .96]],
  student:  [['e', .28, .02, .44, .44], ['t', .50, .48, .06, .98, .94, .98], ['l', .18, .70, .18, .96]],
  face:     [['e', .02, .02, .96, .96], ['E', .28, .34, .11, .11], ['E', .61, .34, .11, .11], ['l', .30, .66, .70, .66]],
  person:   [['e', .30, .02, .40, .40], ['t', .50, .46, .12, .98, .88, .98]],
};

/** Draw one of the outline glyphs above. */
function icon(slide, name, x, y, w, h, color) {
  const lw = Math.max(0.75, Math.min(w, h) * 3.4);
  ICONS[name].forEach((p) => {
    const k = p[0];
    const X = (v) => x + v * w, Y = (v) => y + v * h;
    if (k === 'l') line(slide, X(p[1]), Y(p[2]), X(p[3]), Y(p[4]), { line: { color, width: lw } });
    else if (k === 't') poly(slide, x, y, w, h, [[p[1], p[2]], [p[3], p[4]], [p[5], p[6]]], { fill: { color } });
    else {
      const o = { x: X(p[1]), y: Y(p[2]), w: p[3] * w, h: p[4] * h };
      const type = (k === 'r' || k === 'R') ? 'rect' : 'ellipse';
      if (k === 'r' || k === 'e') shape(slide, type, Object.assign(o, { line: { color, width: lw } }));
      else shape(slide, type, Object.assign(o, { fill: { color } }));
    }
  });
}

/** The recurring ">>" / "<<" double-chevron glyph (two overlapping stroke arrows). */
function chevrons(slide, x, y, w, h, color, dir) {
  const cw = w * 0.63, gap = w * 0.37;
  const pts = [[.20, 0], [.13, .14], [.66, .486], [.09, .864], [.18, 1], [.26, .975], [1, .486], [.31, .03]];
  for (let i = 0; i < 2; i++) {
    poly(slide, x + i * gap, y, cw, h,
      dir === 'left' ? pts.map((p) => [1 - p[0], p[1]]) : pts, { fill: { color } });
  }
}

/** Ring + disc + glyph — the deck's standard circular icon button. */
function iconButton(slide, x, y, size, ringColor, glyph) {
  shape(slide, 'ellipse', { x, y, w: size, h: size, fill: { color: ringColor, transparency: 35 } });
  const d = size * 0.663;
  shape(slide, 'ellipse', { x: x + size * 0.1605, y: y + size * 0.1685, w: d, h: d, fill: { color: PURPLE } });
  if (glyph) {
    const g = size * 0.374;
    icon(slide, glyph, x + size * 0.306, y + size * 0.313, g, g, WHITE);
  }
}

/** Green pill + purple disc + ">>" + caption: the "Explore more" call-to-action. */
function exploreButton(slide, x, y, w, h, discX, discY, disc, textX, textY) {
  shape(slide, 'roundRect', { x, y, w, h, rectRadius: h / 2, fill: { color: PURPLE_D, transparency: 35 } });
  shape(slide, 'ellipse', { x: discX, y: discY, w: disc, h: disc, fill: { color: PURPLE } });
  chevrons(slide, discX + disc * 0.252, discY + disc * 0.234, disc * 0.54, disc * 0.513, WHITE, 'right');
  text(slide, 'Explore more', { x: textX, y: textY, w: 2.229, h: 0.438, fontSize: 20, color: WHITE });
}

/** Master furniture: wordmark top-left, page number bottom-right. */
function chrome(slide, num) {
  text(slide, 'Art Gallery', { x: 0.275, y: 0.208, w: 1.325, h: 0.303, fontSize: 12, fontFace: HEAD });
  shape(slide, 'roundRect', { x: 12.516, y: 6.846, w: 0.475, h: 0.450, rectRadius: 0.225, fill: { color: PURPLE } });
  text(slide, String(num), {
    x: 12.438, y: 6.919, w: 0.630, h: 0.303, fontSize: 12, fontFace: HEAD, color: WHITE, align: 'center',
  });
}

/** Deck title: 44 pt Lora, used on nearly every slide. */
function title(slide, str, x, y, w, h, color, o) {
  text(slide, str, Object.assign({ x, y, w, h, fontSize: TITLE_SZ, fontFace: HEAD, color: color || PURPLE }, o));
}
/** 12 pt body paragraph with the deck's 130% leading. */
function body(slide, str, x, y, w, h, o) {
  text(slide, str, Object.assign({ x, y, w, h, fontSize: 12, lineSpacingMultiple: 1.3 }, o));
}
/** Bold 16 pt purple sub-heading + 12 pt paragraph, the deck's repeating "feature" block. */
function feature(slide, head, hx, hy, hw, bx, by, bw, bh, txt) {
  text(slide, head, { x: hx, y: hy, w: hw, h: 0.417, fontSize: 16, bold: true, color: PURPLE, lineSpacingMultiple: 1.3 });
  body(slide, txt || LEAD2, bx, by, bw, bh);
}
/** Lower half of a process ring, ending in the arrowhead that points at the next step. */
function ringArrowHalf(slide, x, y, w, h, o) {
  const X = (v) => v * w, Y = (v) => v * h;
  shape(slide, 'custGeom', Object.assign({
    x, y, w, h, points: [
      { x: X(.0897), y: Y(1) }, { x: X(0), y: Y(.8208) }, { x: X(.031), y: Y(.8208) }, { x: X(.039), y: Y(.6807) },
      { x: X(.5133), y: Y(0), curve: { type: 'cubic', x1: X(.0841), y1: Y(.2931), x2: X(.2792), y2: Y(.0011) } },
      { x: X(1), y: Y(.8475), curve: { type: 'cubic', x1: X(.7808), y1: Y(-.0013), x2: X(.9985), y2: Y(.3778) } },
      { x: X(.8867), y: Y(.8486) },
      { x: X(.5136), y: Y(.1988), curve: { type: 'cubic', x1: X(.8856), y1: Y(.4885), x2: X(.7187), y2: Y(.1979) } },
      { x: X(.15), y: Y(.7207), curve: { type: 'cubic', x1: X(.3341), y1: Y(.1997), x2: X(.1846), y2: Y(.4235) } },
      { x: X(.1442), y: Y(.8208) }, { x: X(.1793), y: Y(.8208) }, { close: true },
    ],
  }, o));
}

/** Solid pictogram figure (head, shouldered torso, two legs) for the visitor comparison. */
function figurine(slide, x, y, w, h, color) {
  const f = { color };
  shape(slide, 'ellipse', { x: x + w * 0.26, y, w: w * 0.48, h: h * 0.26, fill: f });
  shape(slide, 'roundRect', { x, y: y + h * 0.30, w, h: h * 0.42, rectRadius: w * 0.30, fill: f });
  shape(slide, 'roundRect', { x: x + w * 0.14, y: y + h * 0.55, w: w * 0.28, h: h * 0.45, rectRadius: w * 0.10, fill: f });
  shape(slide, 'roundRect', { x: x + w * 0.58, y: y + h * 0.55, w: w * 0.28, h: h * 0.45, rectRadius: w * 0.10, fill: f });
}

/* ==================================================================== slides */
const slides = [];

/* 1 — cover */
slides.push((s) => {
  navTab(s);
  title(s, 'Art Gallery', 5.044, 0.966, 7.099, 1.447, PURPLE_D, { fontSize: 80, align: 'right' });
  exploreButton(s, 8.938, 5.802, 3.625, 0.812, 11.731, 5.877, 0.663, 9.438, 5.983);
});

/** Green tab with the Art / Gallery / Exhibition nav labels (cover + closing slide). */
function navTab(s) {
  shape(s, 'custGeom', {
    x: 0, y: 0, w: 5.670, h: 0.464, fill: { color: GREEN, transparency: 15 },
    points: [{ x: 0, y: 0 }, { x: 5.670, y: 0 },
      { x: 5.207, y: 0.464, curve: { type: 'cubic', x1: 5.670, y1: 0.256, x2: 5.463, y2: 0.464 } },
      { x: 0, y: 0.464 }, { close: true }],
  });
  [['Art ', 1.033, 0.959], ['Gallery', 2.206, 1.270], ['Exhibition', 3.689, 1.270]].forEach(([t, x, w]) => {
    text(s, t, { x, y: 0.059, w, h: 0.337, fontSize: 14, color: WHITE, align: 'center' });
  });
}

/* 2 — welcome */
slides.push((s) => {
  panelRoundBottom(s, 7.333, 0, 5.042, 4.652, 0.343, { fill: { color: GREEN } });
  title(s, 'Welcome to Our Art Gallery', 0.643, 1.047, 4.753, 1.582);
  body(s, LEAD, 7.990, 1.047, 3.911, 1.388, { color: WHITE });
  iconButton(s, 9.945, 5.261, 1.0, PURPLE_D, null);
  chevrons(s, 10.272, 5.584, 0.358, 0.340, WHITE, 'right');
});

/* 3 — exploring the world of art */
slides.push((s) => {
  title(s, 'Exploring the World of Art', 0.642, 1.816, 4.385, 1.582);
  text(s, 'World of Art', { x: 0.642, y: 3.710, w: 4.765, h: 0.404, bold: true, color: PURPLE });
  body(s, LEAD, 0.642, 4.234, 4.385, 1.126);
  shape(s, 'roundRect', { x: 5.781, y: 1.139, w: 6.781, h: 3.059, rectRadius: 0.208, fill: { color: GREEN, transparency: 10 } });
  body(s, LEAD, 6.667, 2.751, 5.286, 0.863, { color: WHITE, align: 'center' });
  [['palette', 7.101], ['masks', 8.895], ['violin', 10.691]].forEach(([g, x]) => icon(s, g, x, 1.705, 0.755, 0.755, WHITE));
  exploreButton(s, 0.642, 5.791, 3.625, 0.632, 3.635, 5.847, 0.520, 1.143, 5.865);
});

/* 4 — history */
slides.push((s) => {
  shape(s, 'rect', { x: 0, y: 0, w: 6.312, h: 7.5, fill: { color: GREEN } });
  title(s, 'History of our Art Gallery', 0.771, 1.814, 4.385, 1.582, WHITE);
  text(s, 'Our Journey', { x: 0.771, y: 3.832, w: 4.765, h: 0.404, bold: true, color: WHITE });
  body(s, LEAD, 0.771, 4.356, 4.385, 1.126, { color: WHITE });
  text(s, '1998', { x: 9.968, y: 2.938, w: 1.064, h: 0.404, bold: true, color: PURPLE, fontFace: HEAD });
  text(s, 'Year of Founding', { x: 9.968, y: 3.244, w: 2.397, h: 0.303, fontSize: 12, color: PURPLE, fontFace: HEAD });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet orci ' +
       'et consequat. Morbi semper eros vitae tincidunt porta. Mauris euismod. ', 9.968, 3.831, 2.397, 1.651);
  iconButton(s, 9.255, 1.475, 1.0, PURPLE_D, 'flag');
});

/* 5 — timeless charm */
slides.push((s) => {
  shape(s, 'rect', { x: 7.220, y: 0, w: 5.013, h: 7.5, fill: { color: GREEN } });
  title(s, 'A Journey Through Timeless Charm', 0.642, 1.561, 5.896, 1.582);
  text(s, 'History of our art gallery', { x: 7.695, y: 4.590, w: 4.176, h: 0.404, bold: true, color: WHITE });
  body(s, LEAD, 7.695, 5.114, 4.176, 1.126, { color: WHITE });
  iconButton(s, 6.243, 0.508, 0.848, PURPLE_D, null);
  chevrons(s, 6.520, 0.782, 0.304, 0.288, WHITE, 'right');
  iconButton(s, 6.254, 4.994, 0.825, PURPLE_D, null);
  chevrons(s, 6.524, 5.261, 0.296, 0.280, WHITE, 'left');
});

/* 6 — we are private art gallery */
slides.push((s) => {
  shape(s, 'roundRect', { x: 4.158, y: 1.027, w: 9.673, h: 2.795, rectRadius: 0.212, fill: { color: GREEN } });
  title(s, 'We Are Private Art Gallery', 6.304, 1.491, 5.380, 1.582, WHITE);
  icon(s, 'museum', 4.920, 1.785, 0.920, 0.921, WHITE);
  body(s, LEAD, 6.304, 4.126, 5.978, 0.863);
  [['museum', 6.307], ['frame', 7.826], ['easel', 9.344]].forEach(([g, x]) => iconButton(s, x, 5.473, 1.0, PURPLE_D, g));
});

/* 7 — types of galleries */
slides.push((s) => {
  shape(s, 'rect', { x: 0, y: 0.810, w: 5.402, h: 3.289, fill: { color: GREEN } });
  title(s, 'Type of Art Galleries', 0.642, 1.201, 4.385, 1.582, WHITE);
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa..',
    { x: 0.593, y: 2.935, w: 4.086, h: 0.863, fontSize: 12, color: WHITE });
  [['University Gallery', 0.642, 'museum'], ['Private Gallery', 4.722, 'easel'], ['Public Gallery', 8.801, 'frame']]
    .forEach(([label, x, g]) => {
      iconButton(s, x, 5.106, 1.0, PURPLE_D, g);
      feature(s, label, x + 1.226, 4.964, 2.356, x + 1.226, 5.383, 2.647, 0.865);
    });
});

/* 8 — role in society */
slides.push((s) => {
  title(s, 'The Role of Art Galleries in Society', 0.771, 1.134, 5.896, 1.582);
  body(s, LEAD, 7.116, 1.402, 5.651, 0.863);
  [['Cultural Preservation', 3.892, 3.872, 3.844, 'frame'], ['Education Provide', 5.367, 5.304, 3.619, 'board']]
    .forEach(([label, ringY, textY, hw, g]) => {
      iconButton(s, 6.968, ringY, 1.0, PURPLE_D, g);
      feature(s, label, 8.194, textY, hw, 8.194, textY + 0.420, 3.619, 0.601);
    });
});

/* 9 — curating exhibitions */
slides.push((s) => {
  shape(s, 'rect', { x: 0, y: 0, w: 6.667, h: 7.5, fill: { color: GREEN, transparency: 35 } });
  title(s, 'How Art Galleries Curate Exhibitions', 0.686, 2.100, 5.896, 1.582, WHITE);
  body(s, LEAD, 0.686, 3.894, 5.651, 0.863, { color: WHITE });
  [['01', 'Curatorial Process', 1.645, 1.523], ['02', 'Themes and Concepts', 3.322, 3.258],
   ['03', 'Artist Collaborations', 5.046, 4.992]].forEach(([n, label, numY, textY]) => {
    text(s, n, { x: 7.267, y: numY, w: 1.237, h: 1.010, fontSize: 54, bold: true, color: PURPLE, fontFace: HEAD });
    feature(s, label, 8.504, textY, 2.682, 8.504, textY + 0.395, 3.472, 0.626,
      'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo');
  });
});

/* 10 — artist collaborations */
slides.push((s) => {
  shape(s, 'rect', { x: 0, y: 1.642, w: 1.349, h: 3.757, fill: { color: GREEN } });
  title(s, 'Artist Collaborations: Working with artist to showcase their work', 5.783, 1.428, 7.229, 2.322);
  iconButton(s, 4.961, 4.006, 1.0, PURPLE_D, null);
  icon(s, 'quote', 5.293, 4.386, 0.307, 0.241, WHITE);
  body(s, LEAD, 6.258, 4.319, 5.978, 0.863);
  body(s, LEAD, 6.258, 5.399, 5.978, 0.863);
});

/* 11 — special exhibitions */
slides.push((s) => {
  title(s, 'Special Exhibitions and Events', 0.830, 0.998, 4.043, 2.322);
  shape(s, 'rect', { x: 0, y: 3.750, w: 40 / 3, h: 3.750, fill: { color: GREEN } });
  body(s, LEAD, 0.830, 4.208, 4.229, 1.126, { color: WHITE });
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ullamcorper sit amet orci ' +
       'et consequat. Morbi semper eros vitae', 0.830, 5.625, 4.229, 0.863, { color: WHITE });
  [['Monet Exhibition', 6.852, 5.956], ['Manet Exhibition', 10.516, 9.675]].forEach(([label, ringX, textX]) => {
    iconButton(s, ringX, 4.917, 1.0, PURPLE_D, 'easel');
    text(s, label, { x: textX, y: 6.083, w: 2.682, h: 0.417, fontSize: 16, bold: true, color: WHITE, align: 'center' });
  });
});

/* 12 — tickets */
slides.push((s) => {
  title(s, 'How to Get the Tickets', 0.642, 1.732, 4.043, 1.582);
  feature(s, 'Online Sale via Web', 0.642, 3.772, 2.682, 0.642, 4.166, 3.472, 0.626,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo');
  feature(s, 'On the Spot', 0.651, 5.032, 2.682, 0.651, 5.426, 3.472, 0.626,
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo');
  shape(s, 'rect', { x: 5.292, y: 0, w: 4.411, h: 7.5, fill: { color: GREEN, transparency: 35 } });
  [['Online Ticket', '$27K', 1.022, 'laptop'], ['Offine Ticket', '$29K', 3.750, 'receipt']].forEach(([label, price, y, g]) => {
    shape(s, 'roundRect', { x: 6.143, y, w: 6.142, h: 2.047, rectRadius: 0.150, fill: { color: PURPLE_L } });
    shape(s, 'ellipse', { x: 6.587, y: y + 0.631, w: 0.768, h: 0.785, fill: { color: PURPLE } });
    icon(s, g, 6.775, y + 0.815, 0.39, 0.40, WHITE);
    text(s, label, { x: 7.579, y: y + 0.363, w: 2.123, h: 0.404, bold: true, color: WHITE, fontFace: HEAD });
    text(s, [1, 2, 3].map(() => ({
      text: 'Sed ut perspiciatis unde omnis',
      options: { bullet: { characters: '\u2022', indent: 13.5 }, lineSpacingMultiple: 1.4, breakLine: true },
    })), { x: 7.595, y: y + 0.766, w: 2.853, h: 0.919, fontSize: 12, color: WHITE });
    /* price flag: left-pointing pentagon + folded corner, with the amount set vertically */
    shape(s, 'homePlate', { x: 11.181, y: y + 0.371, w: 1.244, h: 1.395, rotate: 180, fill: { color: PURPLE } });
    shape(s, 'rtTriangle', { x: 12.309, y: y + 0.254, w: 0.090, h: 0.144, rotate: 90, flipH: true, fill: { color: 'FADA9E' } });
    text(s, price, { x: 11.293, y: y + 0.816, w: 1.395, h: 0.505, rotate: 90, fontSize: 24, bold: true, color: WHITE, fontFace: HEAD, align: 'center' });
    text(s, '/visit', { x: 11.023, y: y + 0.918, w: 1.395, h: 0.303, rotate: 90, fontSize: 12, color: WHITE, fontFace: HEAD, align: 'center' });
  });
});

/* 13 — membership and support */
slides.push((s) => {
  title(s, 'Art Gallery Membership and Support', 4.571, 0.908, 7.554, 1.582);
  body(s, LEAD, 4.580, 2.775, 5.978, 0.863);
  [['Membership Programs', 4.571, 4.087, 2.682], ['Volunteer Opportunities', 8.546, 4.087, 3.472],
   ['Donations', 4.580, 5.347, 2.682], ['Funding', 8.555, 5.347, 2.682]].forEach(([label, x, y, hw]) => {
    feature(s, label, x, y, hw, x, y + 0.395, 3.472, 0.626,
      'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo');
  });
});

/* 14 — membership programs (pricing table) */
slides.push((s) => {
  shape(s, 'round1Rect', { x: 0, y: 3.049, w: 7.526, h: 4.451, rectRadius: 0.667, flipH: true, fill: { color: GREEN } });
  const hdr = { fontSize: 16, bold: true, color: WHITE, fontFace: HEAD };
  const cell = (t, o) => ({ text: t, options: Object.assign({ align: 'center', valign: 'middle' }, o) });
  const price = (t, sz, fill) => cell(t, { fontSize: sz, bold: true, color: PURPLE, fill: { color: fill } });
  s.addTable([
    [cell('Pack 1', Object.assign({ fill: { color: PURPLE } }, hdr)),
     cell('Pack 2', Object.assign({ fill: { color: PURPLE_L } }, hdr)),
     cell('Pack 3', Object.assign({ fill: { color: PURPLE_D } }, hdr))],
    [price('$129/mo', 14, GREEN_L), price('$162/mo', 12, GREEN_L), price('$183/mo', 12, GREEN_L)],
    [cell('', { fill: { color: GREEN_L } }), cell('', { fill: { color: GREEN_L } }), cell('', { fill: { color: GREEN_L } })],
    [price('Leverage agile ', 12, GREEN_L), price('Leverage agile ', 12, GREEN_L), price('Leverage agile ', 12, GREEN_L)],
    [price('Leverage agile ', 12, GREEN_L), price('Leverage agile ', 12, GREEN_L), price('Leverage agile ', 12, GREEN_L)],
  ], {
    x: 5.631, y: 2.554, w: 6.171, colW: [2.057, 2.057, 2.057], rowH: [0.797, 0.797, 0.797, 0.797, 0.797],
    fontFace: BODY, margin: [3.44, 6.89, 3.44, 6.89], border: { type: 'solid', pt: 1, color: WHITE },
  });
  [[6.517, PURPLE], [8.578, PURPLE_L], [10.639, PURPLE_D]].forEach(([x, c]) => {
    poly(s, x, 4.404, 0.276, 0.276, [[.04, .50], [.20, .34], [.38, .55], [.82, .04], [.97, .17], [.40, .90]], { fill: { color: c } });
  });
  text(s, 'Membership Ticket', { x: 2.410, y: 3.998, w: 2.356, h: 0.774, fontSize: 20, bold: true, color: WHITE, fontFace: HEAD });
  body(s, LEAD3, 2.179, 1.525, 9.071, 0.599, { align: 'center' });
  title(s, 'Membership Programs', 2.938, 0.614, 7.554, 0.841, PURPLE, { align: 'center' });
  icon(s, 'ticket', 1.266, 3.884, 0.822, 0.816, WHITE);
  body(s, LEAD3, 1.266, 5.020, 3.614, 1.388, { color: WHITE });
});

/* 15 — visit our website */
slides.push((s) => {
  shape(s, 'rect', { x: 4.768, y: 0, w: 8.565, h: 7.5, fill: { color: GREEN } });
  browserMock(s, 5.433, 1.110, 7.154, 4.802);
  title(s, 'Visit Our Website', 0.745, 1.451, 3.701, 1.582);
  body(s, LEAD, 0.754, 3.318, 3.460, 1.388);
  exploreButton(s, 2.294, 5.346, 3.625, 0.812, 5.087, 5.421, 0.663, 2.794, 5.527);
});

/** Stand-in for the browser-window artwork on the "Visit Our Website" slide. */
function browserMock(s, x, y, w, h) {
  const BAR = 'BFBFBF', STRIP = 'D9D9D9';
  shape(s, 'rect', { x, y, w, h, fill: { color: WHITE } });
  shape(s, 'rect', { x, y, w, h: 0.227, fill: { color: STRIP } });
  [0.05, 0.18, 0.31].forEach((d) => shape(s, 'ellipse', { x: x + d, y: y + 0.075, w: 0.075, h: 0.075, fill: { color: WHITE } }));
  shape(s, 'rect', { x: x + 0.947, y: y + 0.013, w: 1.353, h: 0.214, fill: { color: BAR } });
  shape(s, 'rect', { x, y: y + 0.227, w, h: 0.246, fill: { color: BAR } });
  [0.09, 0.20, 0.31, 0.42].forEach((d) => shape(s, 'ellipse', { x: x + d, y: y + 0.315, w: 0.08, h: 0.08, fill: { color: STRIP } }));
  shape(s, 'rect', { x: x + 0.860, y: y + 0.243, w: w - 1.220, h: 0.200, fill: { color: WHITE } });
  [0, 0.045, 0.090].forEach((d) => shape(s, 'rect', { x: x + w - 0.32, y: y + 0.290 + d, w: 0.11, h: 0.02, fill: { color: STRIP } }));
  text(s, '[image]', { x, y: y + h / 2 - 0.2, w, h: 0.4, fontSize: 12, color: 'EEEEEE', align: 'center' });
}

/* 16 — curatorial process (dashed path) */
slides.push((s) => {
  const dash = { color: GREEN, width: 3, dashType: 'sysDash' };
  line(s, 0.102, 1.476, 8.596, 1.476, { line: dash });
  line(s, -0.123, 3.863, 8.314, 3.863, { line: dash });
  shape(s, 'arc', { x: 7.403, y: 1.476, w: 2.387, h: 2.387, angleRange: [270, 90], line: dash });
  [['Registration', 2.908, 0.933, 2.076, 2.078, 'clipbrd'], ['Selection', 6.518, 0.933, 5.687, 2.078, 'doc'],
   ['Display Artworks', 1.398, 3.320, 0.567, 4.464, 'easel'], ['Research', 5.008, 3.320, 4.177, 4.464, 'head']]
    .forEach(([label, cx, cy, tx, ty, g]) => {
      iconButton(s, cx, cy, 1.0, PURPLE_D, g);
      text(s, label, { x: tx, y: ty, w: 2.647, h: 0.420, fontSize: 16, bold: true, color: PURPLE, align: 'center', lineSpacingMultiple: 1.3 });
      body(s, LEAD4, tx, ty + 0.419, 2.647, 0.601, { align: 'center' });
    });
  title(s, 'Curratorial Process', 7.663, 4.306, 4.433, 1.582);
  body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin convallis, arcu in aliquam molestie, nibh augue',
    7.663, 5.999, 4.453, 0.626, { color: GREY_80 });
});

/* 17 — visitor analysis (stacked bars) */
slides.push((s) => {
  [[0, 2.968, 6.987, PURPLE], [0, 4.218, 8.111, PURPLE_D], [0, 5.467, 9.235, PURPLE_XD]]
    .forEach(([x, y, w, c]) => shape(s, 'rect', { x, y, w, h: 1.250, fill: { color: c } }));
  [[6.362, 3.592, 2.498, 0.445, PURPLE], [7.593, 4.842, 2.498, 0.472, PURPLE_D], [9.475, 5.466, 1.249, 0.419, PURPLE_XD]]
    .forEach(([x, y, w, r, c]) => shape(s, 'round1Rect', { x, y, w, h: 1.250, rectRadius: r, rotate: 90, flipH: true, fill: { color: c } }));
  [['+84%', 7.047, 3.354], ['+38%', 8.278, 4.604], ['+12%', 9.535, 5.854]].forEach(([t, x, y]) => {
    text(s, t, { x, y, w: 1.111, h: 0.505, fontSize: 24, bold: true, color: GREY_F2, fontFace: HEAD, align: 'center' });
  });
  [['Students', 3.131, 'student', 3.278], ['Civil', 4.391, 'face', 4.589], ['Other', 5.636, 'person', 5.840]]
    .forEach(([label, y, g, cy]) => {
      text(s, label, { x: 2.000, y, w: 2.860, h: 0.370, fontSize: 16, bold: true, color: WHITE, fontFace: HEAD });
      text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin convallis, arcu in aliquam molestie.',
        { x: 2.000, y: y + 0.352, w: 4.316, h: 0.567, fontSize: 12, color: WHITE, lineSpacingMultiple: 1.2 });
      shape(s, 'ellipse', { x: 1.032, y: cy, w: 0.507, h: 0.507, fill: { color: WHITE } });
      icon(s, g, 1.130, cy + 0.098, 0.31, 0.31, PURPLE);
    });
  title(s, 'Visitor Analysis ', 2.861, 0.783, 7.626, 0.774, PURPLE, { align: 'center' });
  body(s, LEAD3, 2.179, 1.634, 9.071, 0.599, { align: 'center' });
});

/* 18 — layer infographic (venn) */
slides.push((s) => {
  [[3.402, 2.974, PURPLE_XD], [1.120, 2.974, PURPLE_L], [2.261, 1.234, PURPLE]]
    .forEach(([x, y, c]) => shape(s, 'ellipse', { x, y, w: 3.264, h: 3.264, fill: { color: c } }));
  title(s, 'Layer Infographic', 7.401, 1.250, 4.016, 1.447);
  const money = { fontSize: 32, bold: true, color: WHITE, lineSpacingMultiple: 1.5 };
  text(s, 'Display', { x: 4.384, y: 4.793, w: 1.605, h: 0.397, fontSize: 16, color: WHITE, align: 'right', lineSpacingMultiple: 1.2 });
  text(s, '$1,289', Object.assign({ x: 3.885, y: 4.992, w: 2.104, h: 0.819, align: 'right' }, money));
  text(s, 'Curration', { x: 1.993, y: 4.777, w: 1.720, h: 0.399, fontSize: 16, color: WHITE, lineSpacingMultiple: 1.2 });
  text(s, '$1,289', Object.assign({ x: 1.993, y: 4.965, w: 2.199, h: 0.819 }, money));
  text(s, 'Research Budget', { x: 2.816, y: 2.800, w: 2.108, h: 0.397, fontSize: 16, color: WHITE, align: 'center', lineSpacingMultiple: 1.2 });
  text(s, '$1,289', Object.assign({ x: 2.816, y: 3.020, w: 2.108, h: 0.819, align: 'center' }, money));
  [[1.680, 3.879, 'doc'], [5.422, 3.859, 'easel'], [3.496, 1.824, 'head']].forEach(([x, y, g]) => {
    shape(s, 'ellipse', { x, y, w: 0.748, h: 0.748, fill: { color: WHITE } });
    icon(s, g, x + 0.187, y + 0.187, 0.374, 0.374, PURPLE);
  });
  [4.004, 4.607, 5.209, 5.812].forEach((y, i) => {
    icon(s, 'checkC', 7.401, y, 0.315, 0.315, GREEN_D);
    body(s, 'Lorem Ipsum is simply dummy text of the printing', 7.869, 3.991 + i * 0.6025, 4.447, 0.340);
  });
  body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ',
    7.401, 2.815, 4.584, 0.601);
});

/* 19 — visitor comparison (pie charts) */
slides.push((s) => {
  const pie = (x, colors, ang) => s.addChart(pptx.ChartType.pie,
    [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [6.5, 3.2] }],
    { x, y: 2.822, w: 5.417, h: 3.456, chartColors: colors, firstSliceAng: ang,
      showLegend: false, showTitle: false, dataBorder: { pt: 1.5, color: WHITE } });
  pie(1.643, [PURPLE, GREY_D9], 0);
  pie(7.148, [GREY_D9, GREEN_D], 133);
  const pct = (n, x, y) => text(s, [
    { text: n, options: { fontSize: 32, bold: true } },
    { text: '%', options: { fontSize: 32, bold: true, baseline: 600 } },
  ], { x, y, w: 2.351, h: 0.801, color: WHITE, align: 'center', lineSpacingMultiple: 1.3 });
  pct('67', 3.252, 4.731);
  text(s, 'Lorem ipsum dolor', { x: 3.239, y: 5.394, w: 2.351, h: 0.320, fontSize: 11, color: WHITE, align: 'center' });
  pct('33', 9.571, 3.708);
  text(s, 'Lorem ipsum dolor', { x: 9.559, y: 4.372, w: 2.351, h: 0.320, fontSize: 11, color: WHITE, align: 'center' });
  [['Student', 1.512, 1.694, PURPLE_D], ['Non Student', 7.104, 7.286, GREEN_D]].forEach(([label, px, tx, c]) => {
    shape(s, 'roundRect', { x: px, y: 3.261, w: 2.214, h: 0.602, rectRadius: 0.301, fill: { color: c } });
    text(s, label, { x: tx, y: 3.305, w: 1.851, h: 0.426, fontSize: 16, bold: true, color: WHITE, fontFace: HEAD, align: 'center', lineSpacingMultiple: 1.3 });
  });
  [1.400, 6.997].forEach((x) => text(s, LEAD2, { x, y: 3.986, w: 2.351, h: 0.736, fontSize: 10, lineSpacingMultiple: 1.3 }));
  body(s, LEAD3, 2.179, 1.616, 9.071, 0.599, { align: 'center' });
  title(s, 'Visitor Comparison', 2.938, 0.704, 7.554, 0.841, PURPLE, { align: 'center' });
});

/* 20 — data table */
slides.push((s) => {
  shape(s, 'roundRect', { x: 2.651, y: 0.639, w: 8.032, h: 4.251, rectRadius: 0.081, fill: { color: WHITE, transparency: 76 } });
  shape(s, 'roundRect', { x: 2.038, y: 0.876, w: 9.258, h: 4.900, rectRadius: 0.093, fill: { color: WHITE, transparency: 52 } });
  const months = ['August', 'September', 'October', 'November', 'December'];
  const grid = [
    ['First Exhibition', 'First Exhibition', 'First Exhibition', 'First Exhibition', 'First Exhibition'],
    ['-', 'Second Exhibition', 'Second Exhibition', 'Second Exhibition', 'Second Exhibition'],
    ['-', '-', 'Third Exhibition', 'Third Exhibition', 'Third Exhibition'],
    ['-', '-', '-', 'Fourth Exhibition', 'Fourth Exhibition'],
    ['-', '-', '-', '', 'Fifth Exhibition'],
  ];
  const rows = [months.map((m, i) => ({
    text: m, options: { fill: { color: i % 2 ? PURPLE_L : PURPLE }, fontSize: 15, bold: true, color: WHITE, fontFace: 'Lato' },
  }))];
  grid.forEach((r, ri) => rows.push(r.map((t) => ({
    text: t, options: { fill: { color: ri % 2 ? WHITE : GREY_F2 }, fontSize: 9, color: INK },
  }))));
  rows.push(months.map(() => ({ text: 'Clear', options: { fill: { color: GREEN_D }, fontSize: 15, color: WHITE } })));
  s.addTable(rows, {
    x: 1.862, y: 3.067, w: 9.609, colW: [1.922, 1.922, 1.922, 1.944, 1.899], rowH: Array(7).fill(0.46),
    align: 'center', valign: 'middle', fontFace: BODY, margin: [1.74, 1.74, 1.74, 1.74],
    border: { type: 'solid', pt: 1, color: 'BBBFC3' },
  });
  body(s, LEAD3, 2.179, 1.616, 9.071, 0.599, { align: 'center' });
  title(s, 'Data Table', 2.938, 0.704, 7.554, 0.841, PURPLE, { align: 'center' });
});

/* 21 — visitor comparison (pictograms) */
slides.push((s) => {
  shape(s, 'rect', { x: 1.273, y: 3.495, w: 10.789, h: 3.018, fill: { color: GREY_F2 } });
  [['Man Visitor', 1.273, GREEN_D, 3.707, GREEN, 2.938, 2.917], ['Woman Visitor', 6.667, PURPLE, 9.102, PURPLE, 8.332, 2.910]]
    .forEach(([label, x, c, mx, mc, tx, ty]) => {
      shape(s, 'rect', { x, y: 2.844, w: 5.394, h: 0.651, fill: { color: c } });
      shape(s, 'flowChartMerge', { x: mx, y: 3.495, w: 0.525, h: 0.212, fill: { color: mc } });
      text(s, label, { x: tx, y: ty, w: 2.064, h: 0.460, fontSize: 18, bold: true, color: WHITE, fontFace: HEAD, align: 'center', lineSpacingMultiple: 1.3 });
    });
  const blurb = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque.';
  body(s, blurb, 1.783, 4.005, 4.372, 0.863, { align: 'center' });
  body(s, blurb, 7.177, 4.005, 4.373, 0.863, { align: 'center' });
  line(s, 6.667, 3.962, 6.667, 5.662, { line: { color: '595959', width: 1, dashType: 'lgDash' } });
  /* pictogram rows: one large figure, then ten small ones showing the filled share */
  const figures = (x0, big, small, step, color, filled) => {
    figurine(s, x0, 5.149, big, 0.781, color);
    for (let i = 0; i < 10; i++) figurine(s, small + i * step, 5.532, 0.135, 0.394, i < filled ? color : GREY_D9);
  };
  figures(1.783, 0.324, 2.437, 0.159, GREEN, 3);
  figures(7.177, 0.290, 7.793, 0.155, PURPLE, 8);
  [['128K', 4.136, '/mo', 5.216], ['729K', 9.474, '/mo', 10.554]].forEach(([n, nx, u, ux]) => {
    text(s, n, { x: nx, y: 5.432, w: 1.240, h: 0.660, fontSize: 28, bold: true, color: INK, fontFace: HEAD, lineSpacingMultiple: 1.3 });
    text(s, u, { x: ux, y: 5.605, w: 0.941, h: 0.420, fontSize: 16, color: INK, fontFace: HEAD, lineSpacingMultiple: 1.3 });
  });
  body(s, LEAD3, 2.179, 1.616, 9.071, 0.599, { align: 'center' });
  title(s, 'Visitor Comparison', 2.938, 0.704, 7.554, 0.841, PURPLE, { align: 'center' });
});

/* 22 — event schedule */
slides.push((s) => {
  const day = (name, num, i) => ({
    text: [{ text: name, options: { fontSize: 10.5, bold: true, breakLine: true } },
      { text: num, options: { fontSize: 28, bold: true } }],
    options: { fill: { color: i % 2 ? PURPLE_L : PURPLE }, color: WHITE, fontFace: HEAD },
  });
  const slot = (time, what, fill) => ({
    text: time ? [{ text: time, options: { bold: true, breakLine: true } }, { text: what, options: {} }] : '',
    options: { fill: { color: fill }, fontSize: 12, color: INK, fontFace: BODY, lineSpacingMultiple: 1.3 },
  });
  const R1 = GREEN_L, R2 = 'F4F5F6';
  s.addTable([
    ['Mon 21', 'Tue 22', 'Wed 23', 'Thu 24', 'Fri 25', 'Sat 26'].map((d, i) => day(d.split(' ')[0], d.split(' ')[1], i)),
    [slot('09:00', 'Creator Plan', R1), slot('08:00', 'Design plan', R1), slot('10:00', 'Expert call', R1),
     slot('09:00', 'Initial Call', R1), slot('09:00', 'Review', R1), slot('10:00', 'Meeting', R1)],
    [slot('11:00', 'Creator Plan', R2), slot('', '', R2), slot('15:00', 'Write title', R2),
     slot('10:00', 'Write Title', R2), slot('18:00', 'Write Title', R2), slot('', '', R2)],
    [slot('14:00', 'Creator Plan', R1), slot('15:00', 'Write Title', R1), slot('', '', R1),
     slot('', '', R1), slot('', '', R1), slot('', '', R1)],
  ], {
    x: 4.000, y: 2.564, w: 8.061, colW: [1.402, 1.285, 1.343, 1.343, 1.343, 1.343],
    rowH: [0.988, 0.981, 1.229, 0.981], margin: [14.3, 0, 3.6, 17.8],
    border: [{ type: 'none' }, { type: 'solid', pt: 1, color: WHITE }, { type: 'none' }, { type: 'none' }],
  });
  shape(s, 'roundRect', { x: 1.373, y: 3.414, w: 2.355, h: 3.307, rectRadius: 0.183, fill: { color: GREEN } });
  icon(s, 'palette', 1.690, 3.952, 0.623, 0.623, WHITE);
  text(s, 'Project Description', { x: 1.531, y: 4.876, w: 2.064, h: 0.381, fontSize: 14, bold: true, color: WHITE, fontFace: HEAD, lineSpacingMultiple: 1.3 });
  body(s, LEAD4, 1.531, 5.410, 2.198, 0.863, { color: WHITE });
  body(s, LEAD3, 2.179, 1.616, 9.071, 0.599, { align: 'center' });
  title(s, 'Event Schedule', 2.938, 0.704, 7.554, 0.841, PURPLE, { align: 'center' });
});

/* 23 — curration exhibition (ring flow) */
slides.push((s) => {
  [[1.333, 3.030, PURPLE_XL, 1.333, 4.224, false], [3.886, 3.023, PURPLE_L, 3.886, 3.030, true],
   [6.449, 3.039, PURPLE, 6.449, 4.234, false], [9.012, 3.033, PURPLE_D, 9.012, 3.039, true]]
    .forEach(([x, y, c, tx, ty, flip]) => {
      shape(s, 'blockArc', { x, y, w: 2.913, h: 2.913, angleRange: [180, 359.7], arcThicknessRatio: 0.233, flipV: flip, fill: { color: c } });
      ringArrowHalf(s, tx, ty, 3.001, 1.709, { fill: { color: c }, rotate: 180, flipV: flip });
    });
  [['Registration', 1.630, 2.280, 2.637, PURPLE_L, 'clipbrd'], ['Selection', 4.228, 4.829, 5.352, PURPLE, 'doc'],
   ['Research', 6.825, 7.450, 2.637, PURPLE_D, 'head'], ['Display', 9.371, 9.959, 5.352, PURPLE_XD, 'easel']]
    .forEach(([label, tx, cx, cy, ringC, g]) => {
      shape(s, 'ellipse', { x: cx, y: cy, w: 1.0, h: 1.0, fill: { color: ringC, transparency: 35 } });
      icon(s, g, cx + 0.29, cy + 0.24, 0.46, 0.49, WHITE);
      text(s, label, { x: tx, y: 3.819, w: 2.244, h: 0.420, fontSize: 16, bold: true, color: PURPLE, align: 'center', lineSpacingMultiple: 1.3 });
      body(s, LEAD4, tx, 4.238, 2.244, 0.863, { align: 'center' });
    });
  body(s, LEAD3, 2.179, 1.616, 9.071, 0.599, { align: 'center' });
  title(s, 'Curration Exhibition', 2.938, 0.704, 7.554, 0.841, PURPLE, { align: 'center' });
});

/* 24 — curration exhibition (arrow funnel) */
slides.push((s) => {
  const CONNECT = [[0, 0], [1, .6547], [.9465, 1], [.0626, .4069]];   // pale diagonal linkers
  const BAR = [[.9427, 1], [1, 0], [0, 0], [.0569, 1]];               // colour bar
  const CAP = [[.7897, 0], [0, 0], [.2321, 1], [1, 1]];               // darker icon cap
  const rows = [
    { y: 2.187, conn: [4.986, 4.083, 1.701], bar: [4.986, 4.494], cap: [4.986, 1.101, PURPLE_XD],
      color: PURPLE_D, num: '04', numX: 6.267, numW: 2.785, icon: ['easel', 5.351, 2.263, 0.452, 0.516],
      label: 'Exhibition Display', labelX: 9.661, labelH: 0.707, labelC: PURPLE_D, para: [1.751, 2.263], pw: 3.551, padj: 0.236 },
    { y: 3.300, conn: [5.399, 3.262, 1.705], bar: [5.395, 3.675], cap: [5.395, 1.072, PURPLE_D],
      color: PURPLE, num: '03', numX: 6.531, numW: 2.538, icon: ['head', 5.765, 3.395, 0.407, 0.476],
      label: 'Research', labelX: 9.240, labelH: 0.404, labelC: PURPLE, para: [2.155, 3.390], pw: 3.952, padj: 0.249 },
    { y: 4.410, conn: [5.803, 2.450, 1.725], bar: [5.803, 2.858], cap: [5.803, 1.037, PURPLE],
      color: PURPLE_L, num: '02', numX: 6.932, numW: 1.639, icon: ['doc', 6.148, 4.556, 0.383, 0.383],
      label: 'selection', labelX: 8.819, labelH: 0.404, labelC: PURPLE_L, para: [2.546, 4.504], pw: 4.346, padj: 0.263 },
    { y: 5.523, conn: null, bar: [6.211, 2.060], cap: [6.211, 1.004, PURPLE_L],
      color: PURPLE_XL, num: '01', numX: 7.277, numW: 1.059, icon: ['clipbrd', 6.531, 5.670, 0.385, 0.408],
      label: 'Registration', labelX: 8.398, labelH: 0.404, labelC: PURPLE_XL, para: [2.954, 5.587], pw: 4.742, padj: 0.270 },
  ];
  rows.forEach((r) => {
    if (r.conn) poly(s, r.conn[0], r.y, r.conn[1], r.conn[2], CONNECT, { fill: { color: GREEN_L } });
  });
  rows.forEach((r) => {
    shape(s, 'parallelogram', { x: 1.525, y: r.y, w: r.pw, h: 0.692, rectRadius: r.padj, flipH: true, fill: { color: GREEN_L } });
    poly(s, r.bar[0], r.y, r.bar[1], 0.692, BAR, { fill: { color: r.color } });
    poly(s, r.cap[0], r.y, r.cap[1], 0.692, CAP, { fill: { color: r.cap[2], transparency: 90 } });
    icon(s, r.icon[0], r.icon[1], r.icon[2], r.icon[3], r.icon[4], WHITE);
    text(s, r.num, { x: r.numX, y: r.y + 0.106, w: r.numW, h: 0.460, fontSize: 16, bold: true, color: WHITE, fontFace: HEAD, lineSpacingMultiple: 1.5 });
    text(s, LEAD4, { x: r.para[0], y: r.para[1], w: 3.055, h: 0.532, fontSize: 12, align: 'right', lineSpacingMultiple: 1.1 });
    text(s, r.label, { x: r.labelX, y: r.y + 0.111, w: 2.199, h: r.labelH, color: r.labelC, fontFace: HEAD });
  });
  title(s, 'Curration Exhibition', 2.938, 0.704, 7.554, 0.841, PURPLE, { align: 'center' });
});

/* 25 — contact */
slides.push((s) => {
  navTab(s);
  title(s, 'Get in Touch with Us!', 1.033, 1.216, 4.636, 3.433, PURPLE_D, { fontSize: 66 });
  shape(s, 'roundRect', { x: 7.467, y: 1.304, w: 4.391, h: 4.903, rectRadius: 0.183, fill: { color: GREEN } });
  body(s, LEAD3, 1.033, 4.840, 4.636, 1.126);
  [[1.879, 'mail', ['artgallery@mail.com', '2829-347'], 8.290, 1.969, 3.289],
   [3.202, 'pin', ['1066 Summit Park Avenue Southfield, MI 48034'], 8.286, 3.362, 3.289],
   [4.526, 'phone', ['091 \u2013 900', '2003 \u2013 9485- 0293'], 8.286, 4.635, 3.151]]
    .forEach(([ringY, g, lines, tx, ty, tw]) => {
      iconButton(s, 6.981, ringY, 1.0, PURPLE_D, g);
      text(s, lines.map((l, i) => ({ text: l, options: { breakLine: i < lines.length - 1 } })),
        { x: tx, y: ty, w: tw, h: 0.776, fontSize: 16, color: WHITE, fontFace: HEAD, lineSpacingMultiple: 1.3 });
    });
});

/* ------------------------------------------------------------------- render */
slides.forEach((build, i) => {
  const slide = pptx.addSlide();
  slide.background = { color: WHITE };
  /* master furniture is drawn first so slide artwork can cover it, as PowerPoint does */
  if (i > 0 && i < 24) chrome(slide, i + 1);
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '10158167-9954-4497-8f8b-af2193b38279_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
