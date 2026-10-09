/**
 * "Building The Future" construction deck - 20 slides, 10 x 5.625 in.
 * Rebuilt with pptxgenjs only. Photographs in the source deck are replaced
 * by flat [image] placeholder blocks of the same shape and position.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- design kit
const FONT = 'Space Grotesk';
const ORANGE = 'FF5E22';
const NAVY = '0F043E';
const WHITE = 'FFFFFF';
const INK = '000000';
const SOFT_INK = '0C0C0C';
const PHOTO = 'F1F1F1'; // stand-in fill for the deck's raster photos
const PHOTO_LABEL = '3A3A3A';
const RING_DARK = '534C74'; // unfilled part of a donut gauge on navy
const RING_LIGHT = 'FB8B61'; // unfilled part of a donut gauge on orange

const NONE = { type: 'none' };
const INSET = [5.4, 5.4, 2.7, 2.7]; // l, r, b, t in points - matches the source

const LOREM_LONG =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod ' +
    'tempor incididunt ut labore et dolore magna aliqua enim ad minim.';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed';
const LOREM_TINY = 'Lorem ipsum dolor amet consec adipiscing elit sed.';

// Reusable outlines, expressed as fractions of the shape's own box.
const PANEL = [[0, 0], [0.753, 0], [1, 1], [0, 1]]; // full-bleed diagonal panel
const BAND = [[0, 0], [0.902, 0], [1, 1], [0, 1]]; // wide banner with slanted end
const WEDGE = [[0, 0], [0.837, 0], [1, 1], [0, 1]]; // half-height banner
const CLIP_TL = [[0.084, 0], [1, 0], [1, 1], [0, 1]]; // bar with clipped top-left
const ARROW = [ // long double-tipped arrow, drawn horizontally
    [1, 0.5], [0.471, 0], [0.435, 0.033], [0.906, 0.478], [0, 0.478],
    [0, 0.525], [0.903, 0.525], [0.435, 0.967], [0.471, 1], [1, 0.5007], [0.999, 0.5],
];

// ------------------------------------------------------------------- helpers
function poly(s, o) {
    const pts = o.pts.map(([px, py]) => ({ x: +(px * o.w).toFixed(4), y: +(py * o.h).toFixed(4) }));
    pts.push({ close: true });
    s.addShape('custGeom', {
        x: o.x, y: o.y, w: o.w, h: o.h, points: pts,
        fill: o.fill ? { color: o.fill } : NONE,
        line: o.line || NONE,
        rotate: o.rot, flipH: o.flipH, flipV: o.flipV,
    });
}

function box(s, kind, o) {
    s.addShape(kind, {
        x: o.x, y: o.y, w: o.w, h: o.h,
        fill: o.fill ? { color: o.fill } : NONE,
        line: o.line || NONE,
        rectRadius: o.r, rotate: o.rot, flipH: o.flipH,
        angleRange: o.angleRange, arcThicknessRatio: o.arcThicknessRatio,
    });
}

/** Rectangle with the top-right and bottom-left corners cut off. */
function snipRect(s, o) {
    const dx = o.snip / o.w;
    const dy = o.snip / o.h;
    poly(s, {
        x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill,
        pts: [[0, 0], [1 - dx, 0], [1, dy], [1, 1], [dx, 1], [0, 1 - dy]],
    });
}

function text(s, body, o) {
    s.addText(body, {
        x: o.x, y: o.y, w: o.w, h: o.h,
        fontFace: FONT, fontSize: o.size, color: o.color || INK, bold: o.bold,
        align: o.align || 'left', valign: o.valign || 'top',
        lineSpacingMultiple: o.ls,
        margin: INSET, isTextBox: true, fit: o.fit || 'resize', wrap: true,
    });
}

/** Two-tone headline: light first half, bold accent second half. */
function headline(s, o) {
    const runs = o.parts.map((part) => ({
        text: part.t,
        options: {
            color: part.c, bold: !!part.b, breakLine: !!part.br, fontSize: part.size || o.size,
        },
    }));
    text(s, runs, { x: o.x, y: o.y, w: o.w, h: o.h, size: o.size, ls: o.ls || 0.9 });
}

function pill(s, o) {
    const h = o.h || 0.358;
    const w = o.w || 1.396;
    box(s, 'roundRect', { x: o.x, y: o.y, w, h, r: h / 2, fill: o.fill, line: o.line });
    text(s, o.label, {
        x: o.x, y: o.y, w, h, size: 12, color: o.labelColor,
        align: 'center', valign: 'middle', fit: 'none',
    });
}

/** Kicker + paragraph + "Learn More" button used down the left of many slides. */
function introBlock(s, o) {
    text(s, o.kicker, { x: 0.4, y: o.y, w: 3.037, h: 0.328, size: 15, color: o.kickerColor });
    text(s, o.body || LOREM_LONG, {
        x: 0.419, y: o.bodyY, w: 3.721, h: 0.744, size: 11, color: o.bodyColor || INK, ls: 1.3,
    });
    pill(s, {
        x: o.btnX || 0.5, y: o.btnY || 4.526,
        fill: o.btnFill || NAVY, label: 'Learn More', labelColor: o.btnLabelColor || WHITE,
    });
}

/**
 * Stand-in for a photograph: a flat block in the picture's outline, marked
 * with the classic "picture" glyph so it reads as [image] rather than as art.
 */
function imageBlock(s, o) {
    if (o.pts) poly(s, { x: o.x, y: o.y, w: o.w, h: o.h, pts: o.pts, fill: PHOTO });
    else box(s, 'rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: PHOTO });
    const gw = o.gw || 0.36;
    const gh = gw * 0.913;
    const gx = (o.cx === undefined ? o.x + o.w / 2 : o.cx) - gw / 2;
    const gy = (o.cy === undefined ? o.y + o.h / 2 : o.cy) - gh / 2;
    const g = (kind, fx, fy, fw, fh, r, fill, rot) => s.addShape(kind, {
        x: gx + fx * gw, y: gy + fy * gh, w: fw * gw, h: fh * gh,
        rectRadius: r ? r * gw : undefined, rotate: rot,
        fill: { color: fill }, line: NONE,
    });
    g('roundRect', 0.00, 0.30, 0.78, 0.68, 0.11, PHOTO_LABEL, 353); // back frame
    g('roundRect', 0.07, 0.37, 0.64, 0.54, 0.07, PHOTO, 353);
    g('roundRect', 0.22, 0.02, 0.76, 0.72, 0.12, PHOTO_LABEL, 0); // front frame
    g('roundRect', 0.29, 0.09, 0.62, 0.58, 0.06, PHOTO, 0);
    g('ellipse', 0.34, 0.15, 0.15, 0.17, 0, PHOTO_LABEL, 0); // sun
    g('triangle', 0.52, 0.19, 0.36, 0.48, 0, PHOTO_LABEL, 0); // big peak
    g('triangle', 0.30, 0.34, 0.30, 0.33, 0, PHOTO_LABEL, 0); // small peak
    g('rect', 0.29, 0.55, 0.62, 0.12, 0, PHOTO_LABEL, 0); // ground
}

// --------------------------------------------------------------------- icons
// Line-art glyphs from the source deck, rebuilt out of native shapes.
// Every icon is drawn inside a 0.423 x 0.381 in cell and scaled from there.

/** Bar-chart glyph: five floating pills over a base rule. */
function iconBars(s, x, y, w, color) {
    const k = w / 0.423;
    const bars = [[0.021, 0.190, 0.085], [0.106, 0.106, 0.169], [0.191, 0.085, 0.106],
    [0.275, 0.127, 0.085], [0.360, 0.000, 0.254]];
    box(s, 'roundRect', { x, y: y + 0.339 * k, w: 0.423 * k, h: 0.042 * k, r: 0.021 * k, fill: color });
    bars.forEach(([bx, by, bh]) => box(s, 'roundRect', {
        x: x + bx * k, y: y + by * k, w: 0.042 * k, h: bh * k, r: 0.021 * k, fill: color,
    }));
}

/** Two hollow arches over a base rule. */
function iconArches(s, x, y, w, color, hole) {
    const k = w / 0.423;
    const t = 0.042 * k;
    box(s, 'roundRect', { x, y: y + 0.338 * k, w: 0.423 * k, h: t, r: t / 2, fill: color });
    [[0.064, 0.084, 0.212], [0.233, 0.000, 0.296]].forEach(([ax, ay, ah]) => {
        const aw = 0.127 * k;
        box(s, 'round2SameRect', { x: x + ax * k, y: y + ay * k, w: aw, h: ah * k, r: aw / 2, fill: color });
        box(s, 'round2SameRect', {
            x: x + ax * k + t, y: y + ay * k + t, w: aw - 2 * t, h: ah * k - t,
            r: (aw - 2 * t) / 2, fill: hole,
        });
    });
}

/** Candlestick glyph: two stems, each wearing a hollow capsule. */
function iconCandles(s, x, y, w, color, hole) {
    const k = w / 0.297;
    const t = 0.033 * k;
    const cw = 0.127 * k;
    [[0, 0, 0.339], [0.170, 0.084, 0.318]].forEach(([cx, cy, ch]) => {
        const bh = 0.56 * ch * k;
        const by = y + cy * k + 0.22 * ch * k;
        box(s, 'roundRect', { x: x + cx * k + cw / 2 - t / 2, y: y + cy * k, w: t, h: ch * k, r: t / 2, fill: color });
        box(s, 'roundRect', { x: x + cx * k, y: by, w: cw, h: bh, r: cw * 0.35, fill: color });
        box(s, 'roundRect', { x: x + cx * k + t, y: by + t, w: cw - 2 * t, h: bh - 2 * t, r: cw * 0.2, fill: hole });
    });
}

/** Three overlapping rings. */
function iconRings(s, x, y, w, color) {
    const k = w / 0.402;
    [[0.191, 0.000, 0.212], [0.106, 0.211, 0.169], [0.000, 0.063, 0.148]].forEach(([rx, ry, d]) => {
        box(s, 'donut', { x: x + rx * k, y: y + ry * k, w: d * k, h: d * k, r: 0.046 * k, fill: color });
    });
}

/** Brand mark: a lozenge outline holding three D shapes. */
function iconLogo(s, x, y, w, h, color, hole) {
    const t = h * 0.21;
    box(s, 'roundRect', { x, y, w, h, r: h / 2, fill: color });
    box(s, 'roundRect', { x: x + t, y: y + t, w: w - 2 * t, h: h - 2 * t, r: (h - 2 * t) / 2, fill: hole });
    for (let i = 0; i < 3; i++) {
        box(s, 'flowChartDelay', {
            x: x + w * 0.155 + i * w * 0.245, y: y + t, w: w * 0.235, h: h - 2 * t,
            fill: null, line: { color, width: h * 8 },
        });
    }
}

/** Testimonial glyph: a message frame plus a small avatar. */
function iconPerson(s, x, y, w, color, hole) {
    const k = w / 0.423;
    box(s, 'roundRect', {
        x, y, w: 0.398 * k, h: 0.33 * k, r: 0.055 * k, fill: null, line: { color, width: 2.6 * k },
    });
    box(s, 'roundRect', { x: x + 0.094 * k, y: y + 0.094 * k, w: 0.212 * k, h: 0.047 * k, r: 0.024 * k, fill: color });
    box(s, 'roundRect', { x: x + 0.094 * k, y: y + 0.188 * k, w: 0.141 * k, h: 0.047 * k, r: 0.024 * k, fill: color });
    // the avatar sits over the frame's lower-right corner, so mask that corner first
    box(s, 'rect', { x: x + 0.285 * k, y: y + 0.2 * k, w: 0.2 * k, h: 0.2 * k, fill: hole });
    box(s, 'ellipse', { x: x + 0.306 * k, y: y + 0.212 * k, w: 0.094 * k, h: 0.094 * k, fill: color });
    box(s, 'blockArc', {
        x: x + 0.285 * k, y: y + 0.26 * k, w: 0.138 * k, h: 0.138 * k, fill: color,
        angleRange: [195, 345], arcThicknessRatio: 0.6,
    });
}

/** Filled circle with a tick inside. */
function iconCheck(s, x, y, d, color, tick) {
    box(s, 'ellipse', { x, y, w: d, h: d, fill: color });
    poly(s, {
        x: x + d * 0.2, y: y + d * 0.28, w: d * 0.6, h: d * 0.44, fill: tick,
        pts: [[0, 0.42], [0.16, 0.24], [0.38, 0.5], [0.85, 0], [1, 0.18], [0.38, 0.9]],
    });
}

/** Map pin: hollow teardrop with a dot in the middle. */
function iconPin(s, x, y, w, h, color, hole) {
    box(s, 'teardrop', { x, y, w, h: h * 0.88, fill: color, rot: 135 });
    box(s, 'teardrop', {
        x: x + w * 0.15, y: y + h * 0.13, w: w * 0.7, h: h * 0.62, fill: hole, rot: 135,
    });
    box(s, 'ellipse', { x: x + w * 0.36, y: y + h * 0.26, w: w * 0.28, h: w * 0.28, fill: color });
}

// -------------------------------------------------------------------- slides
function slide01(p) {
    const s = p.addSlide();
    s.background = { color: WHITE };
    headline(s, {
        x: 0.385, y: 0.914, w: 5.248, h: 1.72, size: 60, ls: 0.8,
        parts: [{ t: 'Building The ', c: INK }, { t: 'Future', c: ORANGE, b: true }],
    });
    text(s, 'Innovative Construction Solutions For Tomorrow',
        { x: 0.421, y: 2.983, w: 4.283, h: 0.682, size: 18, color: ORANGE });
    poly(s, {
        x: 5.296, y: 0, w: 4.704, h: 5.625, fill: ORANGE,
        pts: [[1, 0], [1, 1], [0, 1], [0, 0.575], [0.256, 0.361], [0.256, 0]],
    });
    pill(s, {
        x: 0.5, y: 4.535, w: 2.325, label: 'www.yourwebsite.com', labelColor: ORANGE,
        line: { color: ORANGE, width: 1 },
    });
    iconLogo(s, 5.763, 4.523, 0.666, 0.37, WHITE, ORANGE);
    pill(s, { x: 3.011, y: 4.535, fill: NAVY, label: 'Start Slide', labelColor: WHITE });
    imageBlock(s, {
        x: 7.034, y: 0, w: 2.966, h: 5.625, cx: 8.513, cy: 2.812, gw: 0.617,
        pts: [[0.381, 0], [1, 0], [1, 1], [0, 1], [0, 0.575], [0.381, 0.361]],
    });
}

function slide02(p) {
    const s = p.addSlide();
    imageBlock(s, {
        x: 4.669, y: 0, w: 5.331, h: 5.625, cx: 7.331, cy: 2.812, gw: 0.617,
        pts: [[0, 0], [1, 0], [1, 1], [0.288, 1]],
    });
    poly(s, { x: 0, y: 0, w: 6.218, h: 5.625, pts: PANEL, fill: ORANGE });
    headline(s, {
        x: 0.396, y: 0.553, w: 3.592, h: 1.578, size: 50,
        parts: [{ t: 'About Our ', c: WHITE }, { t: 'Company', c: WHITE, b: true }],
    });
    text(s, LOREM_LONG, { x: 0.427, y: 3.022, w: 3.721, h: 0.744, size: 11, color: WHITE, ls: 1.3 });
    text(s, LOREM_MED, { x: 0.427, y: 3.854, w: 3.083, h: 0.514, size: 11, color: WHITE, ls: 1.3 });
    pill(s, { x: 0.5, y: 4.527, fill: WHITE, label: 'Learn More', labelColor: INK });
    iconLogo(s, 8.834, 0.469, 0.666, 0.37, WHITE, PHOTO);
    text(s, LOREM_MED, { x: 7.303, y: 4.181, w: 2.269, h: 0.748, size: 11, color: WHITE, ls: 1.3, align: 'right' });
}

function slide03(p) {
    const s = p.addSlide();
    headline(s, {
        x: 0.4, y: 0.716, w: 5.912, h: 1.578, size: 50,
        parts: [{ t: 'Building Dreams, ', c: INK }, { t: 'Creating Value', c: ORANGE, b: true }],
    });
    poly(s, { x: 0, y: 2.812, w: 5.0, h: 2.812, pts: WEDGE, fill: ORANGE });
    poly(s, { x: 4.718, y: 2.812, w: 5.282, h: 2.812, pts: WEDGE, fill: NAVY, rot: 180 });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin dum risus sit amet ' +
        'venen atis tristique. Morbi tortor justo.',
        { x: 6.907, y: 1.088, w: 2.693, h: 0.652, size: 9, ls: 1.3 });
    text(s, 'Weather Disrupt', { x: 6.907, y: 0.787, w: 1.852, h: 0.303, size: 14, color: ORANGE });
    [['Our Vision', 0.4, '+122K$', 2.681, 0.4], ['Our Mission', 5.92, '+111M', 2.191, 5.847]]
        .forEach(([title, tx, stat, sw, sx]) => {
            text(s, title, { x: tx, y: 3.161, w: 2.171, h: 0.328, size: 15, color: WHITE });
            text(s, LOREM_LONG, { x: tx, y: 3.489, w: 3.767, h: 0.744, size: 11, color: WHITE, ls: 1.3 });
            text(s, stat, { x: sx, y: 4.398, w: sw, h: 0.631, size: 33, color: WHITE });
        });
}

function slide04(p) {
    const s = p.addSlide();
    headline(s, {
        x: 0.363, y: 0.708, w: 4.26, h: 0.829, size: 50,
        parts: [{ t: 'Core ', c: INK }, { t: 'Values', c: ORANGE, b: true }],
    });
    // Four banners: left ones point right, right ones are mirrored by 180 deg.
    const cards = [
        { title: 'Integrity', fill: ORANGE, x: 0, y: 1.843, w: 5.0, tx: 0.424, bw: 2.461 },
        { title: 'Quality', fill: NAVY, x: 4.851, y: 1.843, w: 5.149, tx: 5.758, bw: 2.483, rot: 180 },
        { title: 'Safety', fill: NAVY, x: 0, y: 3.475, w: 5.0, tx: 0.424, bw: 2.483 },
        { title: 'Innovation', fill: ORANGE, x: 4.851, y: 3.475, w: 5.149, tx: 5.758, bw: 2.483, rot: 180 },
    ];
    cards.forEach((c) => {
        poly(s, { x: c.x, y: c.y, w: c.w, h: 1.4, pts: BAND, fill: c.fill, rot: c.rot });
        text(s, c.title, { x: c.tx, y: c.y + 0.31, w: 2.171, h: 0.328, size: 15, color: WHITE });
        text(s, LOREM_SHORT, { x: c.tx, y: c.y + 0.621, w: c.bw, h: 0.518, size: 11, color: WHITE, ls: 1.3 });
    });
    iconBars(s, 8.911, 2.129, 0.423, WHITE);
    iconArches(s, 8.911, 3.751, 0.423, WHITE, ORANGE);
    iconCandles(s, 3.924, 2.131, 0.297, WHITE, ORANGE);
    iconRings(s, 3.884, 3.777, 0.402, WHITE);
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin dum risus sit amet ' +
        'venen atis tristique. ', { x: 6.882, y: 0.774, w: 2.518, h: 0.652, size: 9, ls: 1.3 });
}

function slide05(p) {
    const s = p.addSlide();
    s.background = { color: ORANGE };
    headline(s, {
        x: 0.4, y: 0.707, w: 5.487, h: 2.328, size: 50,
        parts: [{ t: 'Comprehensive Construction ', c: WHITE }, { t: 'Solutions', c: WHITE, b: true }],
    });
    text(s, LOREM_LONG, { x: 0.438, y: 3.59, w: 3.721, h: 0.744, size: 11, color: WHITE, ls: 1.3 });
    pill(s, { x: 0.51, y: 4.511, fill: WHITE, label: 'Learn More', labelColor: INK });
    poly(s, { x: 6.636, y: 1.909, w: 3.359, h: 3.559, pts: ARROW, fill: WHITE, rot: 45 });
}

function slide06(p) {
    const s = p.addSlide();
    imageBlock(s, {
        x: 0, y: 0, w: 4.891, h: 5.625, cx: 2.442, cy: 2.812, gw: 0.617,
        pts: [[0, 0], [1, 0], [0.653, 1], [0, 1]],
    });
    poly(s, { x: 3.158, y: 0, w: 6.861, h: 5.625, pts: PANEL, fill: ORANGE, flipH: true });
    headline(s, {
        x: 4.9, y: 0.937, w: 4.6, h: 1.578, size: 50,
        parts: [{ t: 'Residential ', c: WHITE }, { t: 'Construction', c: WHITE, b: true }],
    });
    text(s, LOREM_LONG, { x: 4.916, y: 3.012, w: 3.721, h: 0.744, size: 11, color: WHITE, ls: 1.3 });
    text(s, LOREM_MED, { x: 4.916, y: 3.844, w: 3.083, h: 0.514, size: 11, color: WHITE, ls: 1.3 });
    pill(s, { x: 4.989, y: 4.517, fill: WHITE, label: 'Learn More', labelColor: INK });
    poly(s, { x: -0.025, y: 3.436, w: 4.267, h: 1.4, pts: BAND, fill: NAVY });
    text(s, 'Safety', { x: 0.439, y: 3.746, w: 2.171, h: 0.328, size: 15, color: WHITE });
    text(s, LOREM_SHORT, { x: 0.439, y: 4.057, w: 2.483, h: 0.518, size: 11, color: WHITE, ls: 1.3 });
    iconRings(s, 3.193, 3.738, 0.402, WHITE);
}

function slide07(p) {
    const s = p.addSlide();
    imageBlock(s, { x: 0, y: 0, w: 4.31, h: 3.189, cx: 2.153, cy: 1.595, gw: 0.35 });
    poly(s, {
        x: 0, y: 2.325, w: 10.0, h: 3.3, fill: ORANGE,
        pts: [[0, 0], [0.194, 0], [0.292, 0.256], [0.351, 0.256], [0.351, 0.255],
        [1, 0.255], [1, 1], [0.456, 1], [0, 1]],
    });
    headline(s, {
        x: 4.915, y: 0.91, w: 4.655, h: 1.578, size: 50,
        parts: [{ t: 'Commercial ', c: INK }, { t: 'Construction', c: ORANGE, b: true }],
    });
    text(s, LOREM_LONG, { x: 4.915, y: 3.59, w: 3.721, h: 0.744, size: 11, color: WHITE, ls: 1.3 });
    pill(s, { x: 4.987, y: 4.511, fill: WHITE, label: 'Learn More', labelColor: INK });
    text(s, '78,21%', { x: 0.278, y: 3.744, w: 2.642, h: 0.909, size: 50, color: WHITE, align: 'center' });
    text(s, 'Lorem ipsum dolor ', { x: 0.365, y: 4.547, w: 2.134, h: 0.328, size: 15, color: WHITE, align: 'center' });
    text(s, 'Lorem ipsum dolor sit amet, consectetur.',
        { x: 0.428, y: 3.239, w: 1.916, h: 0.455, size: 9, color: WHITE, ls: 1.3 });
}

function slide08(p) {
    const s = p.addSlide();
    headline(s, {
        x: 0.375, y: 0.836, w: 3.417, h: 1.578, size: 50,
        parts: [{ t: 'Industrial ', c: INK }, { t: 'Projects', c: ORANGE, b: true }],
    });
    poly(s, { x: 3.258, y: 0, w: 6.761, h: 5.625, pts: PANEL, fill: ORANGE, flipH: true });
    text(s, LOREM_LONG, { x: 0.427, y: 3.141, w: 3.123, h: 0.978, size: 11, ls: 1.3 });
    const rows = [
        { y: 0.836, ty: 1.018, fill: NAVY, ink: WHITE, pct: '65%', title: 'From Vision to Reality', tw: 2.396 },
        { y: 2.249, ty: 2.427, fill: WHITE, ink: INK, pct: '21%', title: 'Warehouses To Factories', tw: 2.608 },
        { y: 3.663, ty: 3.815, fill: NAVY, ink: WHITE, pct: '43%', title: 'Specialize In Large-scale', tw: 2.396 },
    ];
    rows.forEach((r) => {
        poly(s, { x: 5.031, y: r.y, w: 4.469, h: 1.217, pts: CLIP_TL, fill: r.fill });
        text(s, r.title, { x: 5.76, y: r.ty, w: r.tw, h: 0.303, size: 14, color: r.ink });
        text(s, r.pct, { x: 5.735, y: r.ty + 0.346, w: 1.027, h: 0.53, size: 27, color: r.ink });
        text(s, LOREM_TINY, { x: 6.879, y: r.ty + 0.328, w: 2.505, h: 0.514, size: 11, color: r.ink, ls: 1.3 });
    });
    pill(s, { x: 0.5, y: 4.526, fill: NAVY, label: 'Learn More', labelColor: WHITE });
}

function slide09(p) {
    const s = p.addSlide();
    poly(s, {
        x: 5.047, y: 0, w: 4.953, h: 0.884, fill: NAVY,
        pts: [[0, 0], [1, 0], [1, 0.946], [0.91, 0.946], [0.91, 1], [0.058, 1]],
    });
    headline(s, {
        x: 0.362, y: 0.774, w: 4.963, h: 1.578, size: 50,
        parts: [{ t: 'Renovation ', c: INK, br: true }, { t: '& ', c: INK }, { t: 'Remodeling', c: ORANGE, b: true }],
    });
    poly(s, { x: 5.325, y: 0.837, w: 4.675, h: 1.4, pts: BAND, fill: ORANGE, rot: 180 });
    text(s, 'Quality', { x: 6.159, y: 1.147, w: 2.171, h: 0.328, size: 15, color: WHITE });
    text(s, LOREM_SHORT, { x: 6.159, y: 1.458, w: 2.483, h: 0.518, size: 11, color: WHITE, ls: 1.3 });
    iconBars(s, 9.007, 1.123, 0.423, WHITE);
    introBlock(s, { kicker: 'Provide Renovation Services', kickerColor: ORANGE, y: 3.186, bodyY: 3.559 });
    imageBlock(s, { x: 5.775, y: 2.237, w: 4.225, h: 3.388, cx: 7.885, cy: 3.931, gw: 0.372 });
}

function slide10(p) {
    const s = p.addSlide();
    headline(s, {
        x: 0.382, y: 0.818, w: 9.352, h: 0.833, size: 50,
        parts: [{ t: 'From Design TO ', c: INK }, { t: 'Delivery', c: ORANGE, b: true }],
    });
    const steps = [
        { x: 0.514, fill: ORANGE, snip: 0.389, title: 'Consultation', color: ORANGE, tx: 0.43 },
        { x: 3.028, fill: NAVY, snip: 0.397, title: 'Planning', color: NAVY, tx: 2.952 },
        { x: 5.543, fill: ORANGE, snip: 0.381, title: 'Construction', color: ORANGE, tx: 5.45 },
        { x: 8.058, fill: NAVY, snip: 0.405, title: 'Handover', color: NAVY, tx: 7.974 },
    ];
    steps.forEach((st) => snipRect(s, { x: st.x, y: 2.221, w: 1.446, h: 1.183, snip: st.snip, fill: st.fill }));
    iconBars(s, 5.963, 2.514, 0.622, WHITE);
    iconArches(s, 8.478, 2.514, 0.622, WHITE, NAVY);
    iconCandles(s, 1.027, 2.509, 0.436, WHITE, ORANGE);
    iconRings(s, 3.464, 2.554, 0.591, WHITE);
    [2.391, 4.878, 7.399].forEach((x) => s.addShape('line', {
        x, y: 2.85, w: 0.262, h: 0, line: { color: INK, width: 3, endArrowType: 'triangle' },
    }));
    steps.forEach((st) => {
        text(s, st.title, { x: st.tx, y: 3.692, w: 1.612, h: 0.328, size: 15, color: st.color });
        text(s, 'PLACEHOLDER',
            { x: st.tx + 0.003, y: 4.012, w: 1.756, h: 0.647, size: 9, color: SOFT_INK, ls: 1.3 });
    });
}

function slide11(p) {
    const s = p.addSlide();
    headline(s, {
        x: 0.434, y: 0.738, w: 3.779, h: 1.578, size: 50,
        parts: [{ t: 'Why', c: ORANGE, b: true }, { t: ' Choose Us', c: INK }],
    });
    introBlock(s, { kicker: 'Provide Renovation Services', kickerColor: ORANGE, y: 3.188, bodyY: 3.561 });
    poly(s, { x: 5.0, y: 2.195, w: 5.0, h: 3.43, fill: ORANGE, pts: [[0, 0], [1, 0.439], [1, 1], [0, 1]] });
    text(s, 'Safety', { x: 5.455, y: 3.207, w: 2.171, h: 0.328, size: 15, color: WHITE });
    text(s, '+122K$', { x: 5.372, y: 3.541, w: 2.982, h: 0.985, size: 54, color: WHITE });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin bibendum risus sit ' +
        'amet venen atis tristique. ',
        { x: 5.489, y: 4.448, w: 3.606, h: 0.455, size: 9, color: WHITE, ls: 1.3 });
    imageBlock(s, {
        x: 5.0, y: 0, w: 5.0, h: 3.725, cx: 7.497, cy: 1.863, gw: 0.409,
        pts: [[0, 0], [1, 0], [1, 1], [0, 0.595]],
    });
}

function slide12(p) {
    const s = p.addSlide();
    imageBlock(s, { x: 4.984, y: 0.743, w: 2.159, h: 3.279, cx: 6.061, cy: 2.382, gw: 0.36 });
    imageBlock(s, { x: 7.341, y: 0.743, w: 2.159, h: 3.279, cx: 8.418, cy: 2.382, gw: 0.36 });
    headline(s, {
        x: 0.375, y: 0.719, w: 4.025, h: 1.578, size: 50,
        parts: [{ t: 'Industry ', c: INK }, { t: 'Experience', c: ORANGE, b: true }],
    });
    introBlock(s, { kicker: 'Provide Renovation Services', kickerColor: ORANGE, y: 3.185, bodyY: 3.558 });
    [{ x: 4.984, fill: ORANGE, stat: '80%' }, { x: 7.341, fill: NAVY, stat: '530+' }].forEach((c) => {
        snipRect(s, { x: c.x, y: 3.76, w: 2.159, h: 1.123, snip: 0.187, fill: c.fill });
        text(s, c.stat, { x: c.x + 0.199, y: 3.862, w: 1.109, h: 0.48, size: 24, bold: true, color: WHITE });
        text(s, 'Lorem ipsum dolor sit ametconse adipiscing.',
            { x: c.x + 0.199, y: 4.266, w: 1.76, h: 0.514, size: 11, color: WHITE, ls: 1.3 });
    });
}

function slide13(p) {
    const s = p.addSlide();
    imageBlock(s, {
        x: 4.669, y: 0, w: 5.331, h: 5.625, cx: 7.331, cy: 2.812, gw: 0.617,
        pts: [[0, 0], [1, 0], [1, 1], [0.288, 1]],
    });
    poly(s, { x: 0, y: 0, w: 6.209, h: 5.625, pts: PANEL, fill: ORANGE });
    headline(s, {
        x: 0.4, y: 0.869, w: 3.413, h: 1.578, size: 50,
        parts: [{ t: 'Our ', c: WHITE }, { t: 'Leader', c: WHITE, b: true }],
    });
    text(s, LOREM_LONG, { x: 0.438, y: 3.59, w: 3.721, h: 0.744, size: 11, color: WHITE, ls: 1.3 });
    pill(s, { x: 0.51, y: 4.511, fill: WHITE, label: 'Learn More', labelColor: INK });
    text(s, 'Provide Renovation Services', { x: 0.4, y: 3.262, w: 3.037, h: 0.328, size: 15, color: WHITE });
    poly(s, { x: 5.155, y: 3.476, w: 4.845, h: 1.4, pts: BAND, fill: NAVY, flipH: true });
    text(s, 'Hikam Robinson', { x: 6.039, y: 3.753, w: 2.466, h: 0.328, size: 15, color: WHITE });
    text(s, LOREM_SHORT, { x: 6.039, y: 4.078, w: 2.819, h: 0.518, size: 11, color: WHITE, ls: 1.3 });
    iconPerson(s, 9.078, 3.768, 0.423, WHITE, NAVY);
}

function slide14(p) {
    const s = p.addSlide();
    headline(s, {
        x: 0.383, y: 0.669, w: 3.358, h: 2.328, size: 50,
        parts: [{ t: 'Project', c: ORANGE, b: true }, { t: ' Portfolio Overview', c: INK }],
    });
    introBlock(s, { kicker: 'Provide Renovation Services', kickerColor: ORANGE, y: 3.209, bodyY: 3.558 });
    poly(s, { x: 3.345, y: 0.799, w: 0.543, h: 0.575, pts: ARROW, fill: ORANGE, line: { color: ORANGE, width: 3.25 } });
    imageBlock(s, {
        x: 5.0, y: 0, w: 5.0, h: 3.301, cx: 7.498, cy: 1.651, gw: 0.362,
        pts: [[0, 0], [1, 0], [1, 1], [0, 0.595]],
    });
    imageBlock(s, {
        x: 5.0, y: 2.324, w: 5.0, h: 3.301, cx: 7.498, cy: 3.974, gw: 0.362,
        pts: [[0, 0], [1, 0.405], [1, 1], [0, 1]],
    });
}

function slide15(p) {
    const s = p.addSlide();
    s.background = { color: ORANGE };
    headline(s, {
        x: 0.4, y: 0.707, w: 4.033, h: 2.694, size: 86,
        parts: [{ t: 'Break ', c: WHITE }, { t: 'Slide', c: WHITE, b: true }],
    });
    text(s, LOREM_LONG, { x: 0.438, y: 4.18, w: 3.721, h: 0.744, size: 11, color: WHITE, ls: 1.3 });
    poly(s, { x: 6.636, y: 1.909, w: 3.359, h: 3.559, pts: ARROW, fill: WHITE, rot: 45 });
    text(s, '15 Minutes', { x: 5.0, y: 0.832, w: 2.044, h: 0.48, size: 24, color: WHITE });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipis.',
        { x: 5.0, y: 1.311, w: 2.436, h: 0.518, size: 11, color: WHITE, ls: 1.3 });
}

function slide16(p) {
    const s = p.addSlide();
    poly(s, { x: 3.987, y: 0, w: 6.013, h: 2.875, fill: ORANGE, pts: [[0.144, 0], [1, 0], [1, 1], [0, 1]] });
    headline(s, {
        x: 5.162, y: 0.824, w: 4.338, h: 1.578, size: 50,
        parts: [{ t: 'Safety First, ', c: WHITE }, { t: 'Always', c: WHITE, b: true }],
    });
    const budget = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin dum risus sit amet venen atis';
    [{ x: 0.457, sx: 0.37 }, { x: 2.923, sx: 2.87 }].forEach((c) => {
        text(s, 'Safety Budget', { x: c.x, y: 3.314, w: 1.992, h: 0.303, size: 14, color: ORANGE });
        text(s, budget, { x: c.x, y: 3.616, w: 2.24, h: 0.652, size: 9, ls: 1.3 });
        text(s, '+122K$', { x: c.sx, y: 4.357, w: 1.825, h: 0.631, size: 33, color: ORANGE });
    });
    text(s, '0%', { x: 5.498, y: 3.231, w: 2.409, h: 1.818, size: 104, color: NAVY });
    text(s, 'Incident', { x: 7.882, y: 3.47, w: 1.494, h: 0.303, size: 14, color: NAVY });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin dum risus sit amet',
        { x: 7.882, y: 3.772, w: 1.494, h: 0.849, size: 9, ls: 1.3 });
    imageBlock(s, {
        x: 0, y: 0, w: 4.838, h: 2.875, cx: 2.417, cy: 1.438, gw: 0.316,
        pts: [[0, 0], [1, 0], [1, 0.022], [0.825, 1], [0, 1]],
    });
}

function slide17(p) {
    const s = p.addSlide();
    // Six donut gauges: a track arc plus a value arc, capped by a solid centre.
    const gauges = [
        { x: 0.561, y: 1.371, lx: 0.761, track: RING_LIGHT, hub: ORANGE, value: '12K$', end: 212.94912 },
        { x: 1.908, y: 1.371, lx: 2.081, track: RING_LIGHT, hub: ORANGE, value: '21K$', end: 149.74170 },
        { x: 3.255, y: 1.371, lx: 3.426, track: RING_LIGHT, hub: ORANGE, value: '43K$', end: 173.87268 },
        { x: 5.666, y: 3.499, lx: 5.865, track: RING_DARK, hub: NAVY, value: '12K$', end: 212.94912 },
        { x: 7.012, y: 3.499, lx: 7.186, track: RING_DARK, hub: NAVY, value: '21K$', end: 149.74170 },
        { x: 8.359, y: 3.499, lx: 8.530, track: RING_DARK, hub: NAVY, value: '43K$', end: 173.87268 },
    ];
    function gauge(g) {
        // track wedge, then the white "value" wedge, then a hub disc to hollow it out
        s.addShape('pie', {
            x: g.x, y: g.y, w: 1.061, h: 1.062, angleRange: [125.70785, 270.33263],
            fill: { color: g.track }, line: NONE,
        });
        s.addShape('pie', {
            x: g.x, y: g.y, w: 1.061, h: 1.062, angleRange: [269.72667, g.end],
            fill: { color: WHITE }, line: NONE,
        });
        box(s, 'ellipse', { x: g.x + 0.098, y: g.y + 0.098, w: 0.866, h: 0.866, fill: g.hub });
        text(s, g.value, { x: g.lx, y: g.y + 0.359, w: 0.679, h: 0.328, size: 15, color: WHITE, align: 'center' });
    }
    poly(s, { x: 4.46, y: 2.875, w: 5.54, h: 2.0, pts: BAND, fill: NAVY, flipH: true });
    gauges.slice(3).forEach(gauge);
    poly(s, { x: 0, y: 0.741, w: 5.594, h: 2.0, pts: BAND, fill: ORANGE, rot: 180, flipH: true });
    gauges.slice(0, 3).forEach(gauge);
    [['Chart 4', 5.773, 3.016], ['Chart 5', 7.105, 3.016], ['Chart 6', 8.505, 3.019],
    ['Chart 1', 0.68, 0.882], ['Chart 2', 2.012, 0.882], ['Chart 2', 3.412, 0.885]]
        .forEach(([label, x, y]) => text(s, label, {
            x, y, w: 0.857, h: 0.389, size: 12, color: WHITE, align: 'center', ls: 1.5,
        }));
    headline(s, {
        x: 6.245, y: 0.894, w: 2.716, h: 1.578, size: 50,
        parts: [{ t: 'Budget & ', c: INK }, { t: 'Cost', c: ORANGE, b: true }],
    });
    introBlock(s, { kicker: 'Provide Renovation Services', kickerColor: ORANGE, y: 3.185, bodyY: 3.558 });
}

function slide18(p) {
    const s = p.addSlide();
    snipRect(s, { x: 7.163, y: 0.75, w: 2.382, h: 4.125, snip: 0.449, fill: NAVY });
    snipRect(s, { x: 4.457, y: 0.75, w: 2.382, h: 4.125, snip: 0.449, fill: ORANGE });
    [{ x: 4.457, fill: ORANGE }, { x: 7.163, fill: NAVY }].forEach((card) => {
        text(s, '$650', { x: card.x + 0.406, y: 1.667, w: 1.571, h: 0.631, size: 33, color: WHITE, align: 'center' });
        text(s, 'Package One', { x: card.x + 0.258, y: 1.248, w: 1.867, h: 0.328, size: 15, color: WHITE, align: 'center' });
        [2.404, 2.744, 3.084, 3.423].forEach((ly) => {
            iconCheck(s, card.x + 0.341, ly + 0.092, 0.128, WHITE, card.fill);
            text(s, 'Lorem ipsum dolor', { x: card.x + 0.47, y: ly, w: 1.572, h: 0.285, size: 11, color: WHITE, ls: 1.3 });
        });
        pill(s, { x: card.x + 0.53, y: 4.052, w: 1.324, h: 0.326, fill: WHITE, label: 'Select Now', labelColor: INK });
    });
    headline(s, {
        x: 0.362, y: 0.781, w: 3.157, h: 1.578, size: 50,
        parts: [{ t: 'Pricing ', c: INK }, { t: 'Table', c: ORANGE, b: true }],
    });
    introBlock(s, { kicker: 'Provide Renovation Services', kickerColor: ORANGE, y: 3.185, bodyY: 3.558 });
}

function slide19(p) {
    const s = p.addSlide();
    poly(s, { x: 6.678, y: 0.816, w: 3.322, h: 1.252, fill: NAVY, pts: [[0.111, 0], [1, 0], [1, 1], [0, 1]] });
    [['Phone', '+1234567890', 0.5, 1.403], ['Website', 'https://www.web.com', 1.804, 1.553],
    ['Email', 'mail@mail.com', 3.536, 1.109]].forEach(([label, value, x, w]) => {
        text(s, label, { x, y: 4.314, w, h: 0.283, size: 11, ls: 1.3 });
        text(s, value, { x, y: 4.634, w, h: 0.254, size: 9, ls: 1.3 });
    });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adiiscing elit, sed do eiusmod tempor incidid.. ',
        { x: 7.253, y: 1.181, w: 2.247, h: 0.61, size: 9, color: WHITE, ls: 1.2 });
    pill(s, { x: 4.979, y: 4.526, fill: NAVY, label: 'Learn More', labelColor: WHITE });
    headline(s, {
        x: 0.401, y: 0.904, w: 4.245, h: 2.328, size: 50,
        parts: [{ t: 'Let\u2019s Build The Future ', c: INK }, { t: 'Together!', c: ORANGE, b: true }],
    });
    poly(s, { x: 6.678, y: 3.906, w: 3.322, h: 1.719, fill: ORANGE, pts: [[0.153, 0], [1, 0], [1, 1], [0, 1]] });
    imageBlock(s, {
        x: 4.46, y: 2.068, w: 5.54, h: 1.838, cx: 7.228, cy: 2.987, gw: 0.342,
        pts: [[0.098, 0], [1, 0], [1, 1], [0, 1]],
    });
}

function slide20(p) {
    const s = p.addSlide();
    s.background = { color: ORANGE };
    headline(s, {
        x: 0.4, y: 1.004, w: 8.177, h: 2.556, size: 104,
        parts: [
            { t: 'Thank ', c: WHITE },
            { t: 'You', c: WHITE, b: true, br: true },
            { t: '\t\t  For Watching', c: WHITE, size: 60 },
        ],
    });
    poly(s, { x: 8.017, y: 1.101, w: 1.289, h: 1.366, pts: ARROW, fill: WHITE, rot: 45 });
    const RIBBON = [[0, 0], [0.95, 0], [1, 1], [0, 1]];
    poly(s, { x: 0, y: 4.037, w: 5.0, h: 0.838, pts: RIBBON, fill: NAVY });
    text(s, 'Main Street 233, Building A, 4th Floor, Suite 405, Near City Park, Florida, 32801, United States',
        { x: 0.968, y: 4.228, w: 3.374, h: 0.451, size: 9, color: WHITE, ls: 1.3 });
    iconPin(s, 0.599, 4.274, 0.262, 0.357, WHITE, NAVY);
    poly(s, { x: 5.0, y: 4.037, w: 5.0, h: 0.838, pts: RIBBON, fill: WHITE, rot: 180 });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin dum risus sit amet venen atis',
        { x: 5.533, y: 4.218, w: 3.548, h: 0.455, size: 9, ls: 1.3 });
}

// ---------------------------------------------------------------------- main
function build() {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'DECK_16x9', width: 10, height: 5.625 });
    pptx.layout = 'DECK_16x9';
    pptx.title = 'Building The Future';
    pptx.author = 'Construction Deck';

    [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
        slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
        .forEach((fn) => fn(pptx));

    const out = path.join(__dirname, '15042727-77df-44df-be43-50e8e1743699_grok_final.pptx');
    return pptx.writeFile({ fileName: out }).then(() => console.log('wrote ' + out));
}

build();
