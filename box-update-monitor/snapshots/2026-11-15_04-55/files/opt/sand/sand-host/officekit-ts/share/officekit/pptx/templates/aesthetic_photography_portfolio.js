/**
 * "Aesthetic - Photography Presentation Template" rebuilt with pptxgenjs.
 *
 * Run:  node 113d14f4-c35c-44b5-8bf4-4bdc9e1b74aa_grok_final.js
 * Out:  113d14f4-c35c-44b5-8bf4-4bdc9e1b74aa_grok_final.pptx  (next to this file)
 *
 * Photographs in the original deck are redrawn as flat grey "[image]" plates
 * that keep the original clipping silhouette (arch / leaf / pill / ...).
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const BROWN = '786254'; // theme accent1
const TAN = 'BCA88F'; // theme accent2
const WHITE = 'FFFFFF';
const BODY = '262626'; // tx1 lum 85/15
const PHOTO = 'CDCDCD'; // stand-in for the placeholder photographs
const PHOTO_TXT = 'A0A0A0';

const HEAD = 'Inter'; // display face
const SANS = 'Work Sans'; // body face

const SLIDE_W = 13.333;

/* ---------------------------------------------------------- tiny helpers */

/** Cubic-bezier handle length for a quarter circle of radius 1. */
const K = 0.5523;

/**
 * Outline of a rectangle w x h whose corners are rounded by
 * radii = [topLeft, topRight, bottomRight, bottomLeft] (inches).
 */
function cornerPath(w, h, radii) {
    const [tl, tr, br, bl] = radii;
    const p = [{ x: tl, y: 0, moveTo: true }, { x: w - tr, y: 0 }];
    if (tr) p.push({ x: w, y: tr, curve: { type: 'cubic', x1: w - tr + tr * K, y1: 0, x2: w, y2: tr - tr * K } });
    p.push({ x: w, y: h - br });
    if (br) p.push({ x: w - br, y: h, curve: { type: 'cubic', x1: w, y1: h - br + br * K, x2: w - br + br * K, y2: h } });
    p.push({ x: bl, y: h });
    if (bl) p.push({ x: 0, y: h - bl, curve: { type: 'cubic', x1: bl - bl * K, y1: h, x2: 0, y2: h - bl + bl * K } });
    p.push({ x: 0, y: tl });
    if (tl) p.push({ x: tl, y: 0, curve: { type: 'cubic', x1: 0, y1: tl - tl * K, x2: tl - tl * K, y2: 0 } });
    p.push({ close: true });
    return p;
}

/**
 * Corner-radius recipes, in the order [topLeft, topRight, bottomRight, bottomLeft].
 * The deck only ever uses "half the short side" rounding.
 */
const half = (w, h) => Math.min(w, h) / 2;
const ARCH = (w, h) => [half(w, h), half(w, h), 0, 0]; // "top corners rounded"
const LEAF = (w, h) => [half(w, h), 0, half(w, h), 0]; // "diagonal corners rounded"
const LEAF2 = (w, h) => [0, half(w, h), 0, half(w, h)]; // mirrored diagonal
const PILL = (w, h) => [half(w, h), half(w, h), half(w, h), half(w, h)];
const CORNER_TR = (w, h) => [0, half(w, h), 0, 0]; // "single corner rounded"

/** PowerPoint's default corner inset for the round*Rect presets. */
const sixth = (w, h) => Math.min(w, h) / 6;

/** Rectangle with per-corner rounding, emitted as custom geometry. */
function rounded(slide, o) {
    const radii = typeof o.r === 'function' ? o.r(o.w, o.h) : o.r;
    slide.addShape('custGeom', {
        x: o.x, y: o.y, w: o.w, h: o.h,
        points: cornerPath(o.w, o.h, radii),
        fill: o.fill ? { color: o.fill, transparency: o.transparency } : { type: 'none' },
        line: o.line || { type: 'none' },
        rotate: o.rotate, flipH: o.flipH, flipV: o.flipV, shadow: o.shadow,
    });
}

/**
 * Soft black drop shadow. Returns a fresh object every call because
 * pptxgenjs rewrites the shadow options it is handed (points -> EMU).
 */
function shadow(blur, offset, opacity) {
    return { type: 'outer', blur: blur, offset: offset, angle: 45, color: '000000', opacity: opacity };
}

/** Grey stand-in for a photograph, clipped to the same silhouette. */
function photo(slide, o) {
    rounded(slide, Object.assign({ fill: PHOTO }, o));
    if (o.label !== false) {
        slide.addText('[image]', {
            x: o.x, y: o.y + o.h / 2 - 0.18, w: o.w, h: 0.36,
            align: 'center', valign: 'middle', fontFace: SANS, fontSize: 11, color: PHOTO_TXT,
        });
    }
}

/** Body copy: 10.5pt Work Sans, justified, 1.5 line spacing. */
function body(slide, str, o) {
    slide.addText(str, Object.assign({
        fontFace: SANS, fontSize: 10.5, color: BODY,
        align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
    }, o));
}

/** Small bold label, e.g. "1. Vision". */
function label(slide, str, o) {
    slide.addText(str, Object.assign({
        fontFace: SANS, fontSize: 14, bold: true, color: BROWN, valign: 'top', wrap: false,
    }, o));
}

/** Two-tone section heading: tan first half, brown second half. */
function heading(slide, first, second, o) {
    slide.addText([
        { text: first, options: { color: TAN } },
        { text: second, options: { color: BROWN } },
    ], Object.assign({
        fontFace: HEAD, fontSize: 40, bold: true, valign: 'top',
    }, o));
}

/** Pill button with centred bold caption. */
function pill(slide, str, o) {
    slide.addShape('roundRect', {
        x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: half(o.w, o.h),
        fill: { color: o.fill || BROWN },
        line: { type: 'none' },
    });
    if (str) {
        slide.addText(str, {
            x: o.x, y: o.y, w: o.w, h: o.h, align: 'center', valign: 'middle',
            fontFace: SANS, fontSize: o.fontSize || 14, bold: true,
            color: o.color || WHITE, charSpacing: o.charSpacing,
        });
    }
}

/** The three little diagonal-corner squares used as a decorative accent. */
function dots(slide, x, y, dir) {
    const w = 0.179, h = 0.147, step = 0.505;
    for (let i = 0; i < 3; i++) {
        rounded(slide, {
            x: dir === 'h' ? x + i * step : x,
            y: dir === 'h' ? y : y + i * step,
            w: w, h: h, r: LEAF, fill: BROWN, rotate: dir === 'h' ? 270 : 0,
        });
    }
}

/** Top navigation strip that repeats on every slide. */
function navBar(slide) {
    // Brand mark: two interlocking bars, drawn as an "n" and a "u".
    slide.addShape('custGeom', {
        x: 1.158, y: 0.508, w: 0.129, h: 0.18, fill: { color: BROWN }, line: { type: 'none' },
        points: [
            { x: 0.0000, y: 0.0000, moveTo: true }, { x: 0.0470, y: 0.0000 },
            { x: 0.0730, y: 0.0630, curve: { type: 'cubic', x1: 0.0730, y1: 0.0000, x2: 0.0730, y2: 0.0290 } },
            { x: 0.0730, y: 0.1260 }, { x: 0.0210, y: 0.1260 }, { x: 0.0210, y: 0.1800 },
            { x: 0.0000, y: 0.1800 }, { close: true },
            { x: 0.0930, y: 0.0500, moveTo: true }, { x: 0.1290, y: 0.0500 },
            { x: 0.1290, y: 0.1170 },
            { x: 0.0820, y: 0.1800, curve: { type: 'cubic', x1: 0.1290, y1: 0.1520, x2: 0.1090, y2: 0.1800 } },
            { x: 0.0380, y: 0.1800 }, { x: 0.0380, y: 0.1550 }, { x: 0.0820, y: 0.1550 },
            { x: 0.0930, y: 0.1170, curve: { type: 'cubic', x1: 0.0930, y1: 0.1550, x2: 0.0930, y2: 0.1400 } },
            { close: true },
        ],
    });
    const items = [
        { t: 'Home', x: 1.311, w: 0.623, b: true },
        { t: 'AESTHETIC', x: 6.151, w: 1.031, b: true },
        { t: 'Collection', x: 10.130, w: 0.928, b: false },
        { t: 'Gallery', x: 11.288, w: 0.703, b: false },
    ];
    items.forEach(it => slide.addText(it.t, {
        x: it.x, y: 0.459, w: it.w, h: 0.278, wrap: false, valign: 'top',
        fontFace: SANS, fontSize: 10.5, bold: it.b, color: BROWN,
    }));
    // Chevron next to "Gallery".
    slide.addShape('custGeom', {
        x: 11.991, y: 0.575, w: 0.185, h: 0.099, fill: { color: BROWN }, line: { type: 'none' },
        points: [
            { x: 0.023, y: 0.004, moveTo: true }, { x: 0.0925, y: 0.069 }, { x: 0.163, y: 0.003 },
            { x: 0.185, y: 0.021 }, { x: 0.0925, y: 0.099 }, { x: 0.000, y: 0.021 }, { close: true },
        ],
    });
}

const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut';
const LOREM_PROJ = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor sed do '
    + 'PLACEHOLDER';

/* ------------------------------------------------------------ the slides */

function slide01(pres) {
    const s = pres.addSlide();
    s.addShape('round1Rect', { x: 0, y: 3.75, w: 2.979, h: 3.75, rectRadius: 1.4895, fill: { color: BROWN }, line: { type: 'none' } });
    rounded(s, { x: 1.474, y: 1.200, w: 3.280, h: 4.192, r: ARCH, fill: TAN });
    navBar(s);
    photo(s, { x: 1.563, y: 1.300, w: 3.102, h: 3.992, r: ARCH });
    s.addText('AEST', { x: 5.742, y: 1.516, w: 5.326, h: 2.508, wrap: false, valign: 'top', fontFace: HEAD, fontSize: 143, color: BROWN });
    s.addText('HETIC', { x: 5.656, y: 3.388, w: 6.339, h: 2.508, wrap: false, valign: 'top', fontFace: HEAD, fontSize: 143, bold: true, color: TAN });
    s.addText('Photography Presentation Template', {
        x: 5.742, y: 5.821, w: 4.649, h: 0.404, wrap: false, valign: 'top', fontFace: SANS, fontSize: 18, color: BROWN,
    });
    s.addShape('line', { x: 10.594, y: 6.065, w: 2.739, h: 0, line: { color: '3C312A', width: 1 } });
    dots(s, 3.472, 5.867, 'v');
    s.addShape('round1Rect', { x: 0.467, y: 4.278, w: 2.027, h: 2.747, rectRadius: 1.0135, fill: { color: TAN }, line: { type: 'none' } });
    photo(s, { x: 0.654, y: 4.473, w: 1.653, h: 2.357, r: CORNER_TR });
}

function slide02(pres) {
    const s = pres.addSlide();
    navBar(s);
    heading(s, 'Welcome To ', 'Aesthetic', { x: 7.774, y: 1.239, w: 3.705, h: 1.447 });
    body(s, LOREM_SHORT, { x: 7.754, y: 3.091, w: 4.422, h: 0.601 });
    label(s, 'About Our Planning', { x: 7.754, y: 4.211, w: 2.067, h: 0.337 });
    body(s, LOREM_SHORT + 'PLACEHOLDER'
        + 'fugiat nulla pariatur. Excepteur sint occaecat. sed do eiusmod tempor incididunt ut labore et',
        { x: 7.774, y: 4.728, w: 4.385, h: 1.397 });
    dots(s, 1.142, 6.823, 'h');
    photo(s, { x: 3.946, y: 3.847, w: 2.721, h: 2.263, r: LEAF });
    photo(s, { x: 3.946, y: 1.389, w: 2.721, h: 2.263, r: LEAF2 });
    photo(s, { x: 1.158, y: 1.389, w: 2.621, h: 4.716, r: LEAF });
}

function slide03(pres) {
    const s = pres.addSlide();
    const LOREM = 'Lorem ipsum dolor sit amet, consectetur adip iscing elit, sed do';
    rounded(s, { x: 7.524, y: 1.401, w: 2.621, h: 4.716, r: LEAF, fill: TAN });
    photo(s, { x: 7.609, y: 1.497, w: 2.446, h: 4.524, r: LEAF });
    heading(s, 'Our Vision ', 'And Mission', { x: 1.047, y: 1.237, w: 3.942, h: 1.447 });
    [{ x: 1.047, n: '1' }, { x: 3.919, n: '2' }].forEach(col => {
        label(s, col.n + '. Vision', { x: col.x, y: 3.090, w: 1.2, h: 0.337 });
        body(s, LOREM, { x: col.x, y: 3.520, w: 2.141, h: 0.867 });
        label(s, col.n + '. Mission', { x: col.x, y: 4.906, w: 1.2, h: 0.337 });
        body(s, LOREM, { x: col.x, y: 5.337, w: 2.141, h: 0.867 });
    });
    navBar(s);
    rounded(s, { x: 9.637, y: 1.401, w: 2.446, h: 2.443, r: LEAF, fill: TAN });
    rounded(s, { x: 9.637, y: 4.235, w: 2.446, h: 2.443, r: LEAF, fill: TAN });
    dots(s, 1.159, 6.816, 'h');
    photo(s, { x: 9.712, y: 4.315, w: 2.446, h: 2.443, r: LEAF });
    photo(s, { x: 9.723, y: 1.482, w: 2.446, h: 2.443, r: LEAF });
}

function slide04(pres) {
    const s = pres.addSlide();
    rounded(s, { x: 1.176, y: 2.340, w: 3.280, h: 3.767, r: ARCH, fill: TAN });
    navBar(s);
    label(s, 'About Our Facilities', { x: 7.368, y: 1.355, w: 2.136, h: 0.337 });
    body(s, LOREM_SHORT + ' ut labore et dolore reprehenderit in voluptate velit esse cillum dolore. '
        + 'dolore reprehenderit in voluptate velit esse cillum dolore.', { x: 7.368, y: 1.785, w: 4.934, h: 1.132 });
    label(s, '1.Facilities', { x: 7.368, y: 3.522, w: 1.233, h: 0.337 });
    body(s, LOREM_MED, { x: 7.368, y: 3.952, w: 4.934, h: 0.601 });
    label(s, '2.Facilities', { x: 7.368, y: 5.181, w: 1.266, h: 0.337 });
    body(s, LOREM_MED, { x: 7.368, y: 5.611, w: 4.934, h: 0.601 });
    heading(s, 'Our Best ', 'Facilities', { x: 1.058, y: 1.258, w: 5.353, h: 0.774 });
    dots(s, 1.142, 6.823, 'h');
    photo(s, { x: 1.265, y: 2.440, w: 3.102, h: 3.579, r: ARCH });
    rounded(s, { x: 3.781, y: 2.795, w: 2.131, h: 2.581, r: ARCH, fill: TAN });
    photo(s, { x: 3.849, y: 2.876, w: 1.996, h: 2.426, r: ARCH });
}

function slide05(pres) {
    const s = pres.addSlide();
    navBar(s);
    pill(s, '1. Project Title', { x: 1.189, y: 2.619, w: 1.981, h: 0.398 });
    body(s, LOREM_PROJ, { x: 1.073, y: 3.292, w: 4.847, h: 0.867 });
    pill(s, '2. Project Title', { x: 1.191, y: 4.693, w: 1.981, h: 0.398 });
    body(s, LOREM_PROJ, { x: 1.073, y: 5.345, w: 4.847, h: 0.867 });
    heading(s, 'Our ', 'Project', { x: 1.058, y: 1.258, w: 3.953, h: 0.774 });
    dots(s, 7.381, 4.954, 'v');
    dots(s, 12.014, 1.384, 'v');
    // Three rounded photo capsules, each on a brown backing plate.
    const caps = [
        { bx: 7.413, by: 1.400, bw: 1.640, bh: 2.350, px: 7.471, py: 1.482, pw: 1.524, ph: 2.185 },
        { bx: 10.575, by: 3.758, bw: 1.640, bh: 2.350, px: 10.632, py: 3.821, pw: 1.524, ph: 2.225 },
        { bx: 8.627, by: 1.392, bw: 2.520, bh: 4.708, px: 8.710, py: 1.475, pw: 2.354, ph: 4.549 },
    ];
    caps.forEach(c => {
        rounded(s, { x: c.bx, y: c.by, w: c.bw, h: c.bh, r: PILL, fill: BROWN });
        photo(s, { x: c.px, y: c.py, w: c.pw, h: c.ph, r: PILL });
    });
}

function slide06(pres) {
    const s = pres.addSlide();
    navBar(s);
    pill(s, 'Chidi Eze', { x: 1.158, y: 5.489, w: 2.379, h: 0.611 });
    pill(s, 'Drew Feig', { x: 3.870, y: 5.489, w: 2.379, h: 0.611 });
    heading(s, 'Meet Our ', 'Team', { x: 7.329, y: 1.221, w: 4.560, h: 0.774 });
    body(s, LOREM_PROJ + LOREM_PROJ + 'Lorem ipsum dolor sit', { x: 7.351, y: 2.476, w: 4.847, h: 1.662 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor sed do eiusmod sed do',
        { x: 7.345, y: 4.639, w: 4.847, h: 0.601 });
    rounded(s, { x: 7.451, y: 6.061, w: 2.535, h: 0.100, r: PILL, fill: TAN });
    rounded(s, { x: 9.809, y: 6.062, w: 2.367, h: 0.100, r: PILL, fill: BROWN });
    photo(s, { x: 3.870, y: 1.400, w: 2.379, h: 3.804, r: ARCH });
    photo(s, { x: 1.158, y: 1.400, w: 2.379, h: 3.804, r: ARCH });
}

/** The three tiny glyphs down the left edge of slide 7: wallet, clipboard, chart. */
function serviceIcon(slide, kind, x, y) {
    if (kind === 'wallet') {
        slide.addShape('rect', { x: x + 0.085, y: y + 0.030, w: 0.176, h: 0.170, fill: { color: BROWN }, line: { type: 'none' } });
        slide.addShape('rect', { x: x + 0.155, y: y + 0.065, w: 0.106, h: 0.020, fill: { color: WHITE }, line: { type: 'none' } });
        slide.addShape('rect', { x: x + 0.155, y: y + 0.110, w: 0.106, h: 0.020, fill: { color: WHITE }, line: { type: 'none' } });
        slide.addShape('parallelogram', {
            x: x - 0.005, y: y + 0.010, w: 0.190, h: 0.233, rotate: 6,
            fill: { color: BROWN }, line: { type: 'none' },
        });
        slide.addShape('ellipse', { x: x + 0.120, y: y + 0.120, w: 0.030, h: 0.042, fill: { color: WHITE }, line: { type: 'none' } });
    } else if (kind === 'clipboard') {
        slide.addShape('rect', { x: x, y: y + 0.048, w: 0.237, h: 0.196, fill: { color: BROWN }, line: { type: 'none' } });
        slide.addShape('rect', { x: x + 0.040, y: y + 0.020, w: 0.157, h: 0.045, fill: { color: BROWN }, line: { type: 'none' } });
        slide.addShape('rect', { x: x + 0.040, y: y + 0.048, w: 0.157, h: 0.030, fill: { color: WHITE }, line: { type: 'none' } });
        slide.addShape('ellipse', { x: x + 0.098, y: y + 0.000, w: 0.044, h: 0.044, fill: { color: BROWN }, line: { type: 'none' } });
        slide.addShape('custGeom', {
            x: x + 0.040, y: y + 0.105, w: 0.160, h: 0.105, fill: { color: WHITE }, line: { type: 'none' },
            points: [{ x: 0.145, y: 0.000, moveTo: true }, { x: 0.160, y: 0.030 }, { x: 0.062, y: 0.105 },
                { x: 0.000, y: 0.055 }, { x: 0.018, y: 0.028 }, { x: 0.066, y: 0.062 }, { close: true }],
        });
    } else {
        [[0.017, 0.185, 0.048], [0.087, 0.152, 0.081], [0.155, 0.120, 0.113], [0.225, 0.088, 0.145]]
            .forEach(b => slide.addShape('rect', { x: x + b[0], y: y + b[1], w: 0.052, h: b[2], fill: { color: BROWN }, line: { type: 'none' } }));
        // Zig-zag trend line rising to the right.
        slide.addShape('custGeom', {
            x: x, y: y, w: 0.278, h: 0.147, fill: { color: BROWN }, line: { type: 'none' },
            points: [{ x: 0.000, y: 0.128, moveTo: true }, { x: 0.107, y: 0.045 }, { x: 0.148, y: 0.078 },
                { x: 0.238, y: 0.010 }, { x: 0.278, y: 0.010 }, { x: 0.278, y: 0.040 }, { x: 0.250, y: 0.040 },
                { x: 0.150, y: 0.118 }, { x: 0.110, y: 0.085 }, { x: 0.020, y: 0.147 }, { close: true }],
        });
    }
}

function slide07(pres) {
    const s = pres.addSlide();
    const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.';
    rounded(s, { x: 8.135, y: 1.392, w: 2.621, h: 4.716, r: LEAF, fill: TAN });
    photo(s, { x: 8.220, y: 1.488, w: 2.446, h: 4.524, r: LEAF });
    navBar(s);
    const rows = [
        { y: 2.743, btn: 2.316, icon: 2.408, kind: 'wallet', n: '1' },
        { y: 4.150, btn: 3.751, icon: 3.812, kind: 'clipboard', n: '2' },
        { y: 5.556, btn: 5.185, icon: 5.232, kind: 'chart', n: '3' },
    ];
    rows.forEach(r => body(s, LOREM, { x: 1.603, y: r.y, w: 4.189, h: 0.601 }));
    rows.forEach(r => serviceIcon(s, r.kind, 1.170, r.icon));
    rows.forEach(r => pill(s, r.n + '. Service', { x: 1.653, y: r.btn, w: 1.564, h: 0.398 }));
    heading(s, 'Our Best ', 'Service', { x: 1.045, y: 1.222, w: 4.987, h: 0.774 });
    rounded(s, { x: 9.643, y: 1.532, w: 2.446, h: 2.443, r: LEAF, fill: TAN });
    dots(s, 12.015, 4.934, 'v');
    dots(s, 6.661, 1.392, 'v');
    rounded(s, { x: 6.668, y: 3.418, w: 2.446, h: 2.443, r: LEAF, fill: TAN });
    photo(s, { x: 9.728, y: 1.613, w: 2.446, h: 2.443, r: LEAF });
    photo(s, { x: 6.728, y: 3.494, w: 2.446, h: 2.443, r: LEAF });
}

function slide08(pres) {
    const s = pres.addSlide();
    photo(s, { x: 1.771, y: 1.341, w: 2.320, h: 3.586, r: PILL });
    navBar(s);
    heading(s, 'Our ', 'Portofolio', { x: 7.329, y: 1.232, w: 4.662, h: 0.774 });
    body(s, LOREM_PROJ + 'Lorem', { x: 7.339, y: 2.276, w: 4.847, h: 0.867 });
    s.addShape('round1Rect', { x: 0, y: 3.75, w: 2.747, h: 3.75, rectRadius: 1.3735, fill: { color: BROWN }, line: { type: 'none' } });
    rounded(s, { x: 3.116, y: 3.750, w: 2.747, h: 3.750, r: ARCH, fill: BROWN });
    dots(s, 5.680, 1.428, 'v');
    photo(s, { x: 6.667, y: 3.750, w: 5.509, h: 3.175, r: LEAF });
    photo(s, { x: 3.210, y: 3.848, w: 2.559, h: 3.568, r: ARCH });
    photo(s, { x: 0.073, y: 3.846, w: 2.581, h: 3.570, r: CORNER_TR });
}

/** Section title that heads every "Infographic Section" slide. */
function infographicTitle(slide) {
    heading(slide, 'Infographic ', 'Section', { x: 2.892, y: 1.222, w: 7.567, h: 0.774, align: 'center' });
}

function slide09(pres) {
    const s = pres.addSlide();
    s.addShape('round1Rect', { x: 1.354, y: 2.992, w: 2.941, h: 2.941, rectRadius: sixth(2.941, 2.941), fill: { color: BROWN }, line: { type: 'none' } });
    s.addShape('round1Rect', { x: 4.625, y: 2.421, w: 4.084, h: 4.084, rectRadius: sixth(4.084, 4.084), fill: { color: TAN }, line: { type: 'none' } });
    const ctr = { align: 'center', fontFace: SANS, color: WHITE, valign: 'top', lineSpacingMultiple: 1.5 };
    // Featured "Premium" plan.
    s.addText('$75.00', Object.assign({ x: 5.014, y: 2.972, w: 3.347, h: 1.319, fontSize: 54, bold: true }, ctr));
    s.addText('Premium', Object.assign({ x: 4.999, y: 2.745, w: 3.347, h: 0.642, fontSize: 24, bold: true }, ctr));
    s.addShape('line', { x: 5.145, y: 4.442, w: 3.044, h: 0, line: { color: WHITE, width: 1 } });
    s.addText('Lorem ipsum dolor sit amet, consec tetuer adipiscing elit. ',
        Object.assign({ x: 5.145, y: 4.690, w: 3.044, h: 0.603, fontSize: 10.5 }, ctr));
    pill(s, 'ORDER NOW', { x: 5.145, y: 5.516, w: 3.044, h: 0.552, fill: WHITE, color: TAN, fontSize: 12, charSpacing: 3 });
    // Side plans.
    const side = [
        { name: 'Reguler', price: '$50.00', x: 1.354, bx: 1.735, tx: 1.541, px: 1.872, by: 5.180, ly: 4.224, py: 4.394, ty: 3.182, qy: 3.477 },
        { name: 'Platinum', price: '$100.00', x: 9.038, bx: 9.418, tx: 9.224, px: 9.556, by: 5.159, ly: 4.204, py: 4.374, ty: 3.161, qy: 3.457 },
    ];
    side.forEach((p, i) => {
        if (i === 1) s.addShape('round1Rect', { x: p.x, y: 2.972, w: 2.941, h: 2.941, rectRadius: sixth(2.941, 2.941), fill: { color: BROWN }, line: { type: 'none' } });
        pill(s, 'ORDER NOW', { x: p.bx, y: p.by, w: 2.160, h: 0.418, fill: WHITE, color: BROWN, fontSize: 10.5, charSpacing: 3 });
        s.addText(p.name, Object.assign({ x: p.tx, y: p.ty, w: 2.568, h: 0.462, fontSize: 16, bold: true }, ctr));
        s.addText(p.price, Object.assign({ x: p.px, y: p.qy, w: 1.736, h: 0.642, fontSize: 24, bold: true }, ctr));
        s.addShape('line', { x: p.bx, y: p.ly, w: 2.160, h: 0, line: { color: WHITE, width: 1 } });
        s.addText('Lorem ipsum dolor sit am et, consec tetuer',
            Object.assign({ x: p.bx, y: p.py, w: 2.160, h: 0.603, fontSize: 10.5 }, ctr));
    });
    navBar(s);
    infographicTitle(s);
}

function slide10(pres) {
    const s = pres.addSlide();
    const cards = [
        { x: 1.177, fill: BROWN, n: '01.' }, { x: 4.087, fill: TAN, n: '02.' },
        { x: 6.997, fill: BROWN, n: '03.' }, { x: 9.906, fill: TAN, n: '04.' },
    ];
    cards.forEach(c => {
        s.addShape('round1Rect', { x: c.x, y: 3.75, w: 2.281, h: 2.375, rectRadius: sixth(2.281, 2.375), fill: { color: c.fill }, line: { type: 'none' } });
        s.addShape('rect', { x: c.x, y: 3.75, w: 0.269, h: 0.269, fill: { color: WHITE, transparency: 80 }, line: { type: 'none' } });
    });
    const ctr = { align: 'center', fontFace: SANS, color: WHITE, valign: 'top' };
    cards.forEach(c => {
        s.addText('Title Here', Object.assign({ x: c.x + 0.309, y: 4.765, w: 1.662, h: 0.303, fontSize: 12, bold: true }, ctr));
        s.addText('Lorem ipsum dolor sita met, elit. ',
            Object.assign({ x: c.x + 0.309, y: 5.072, w: 1.662, h: 0.627, fontSize: 11, lineSpacingMultiple: 1.5 }, ctr));
    });
    body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue ipsum dolor '
        + 'sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere porttitor congue massa. ',
        { x: 2.195, y: 2.541, w: 8.960, h: 0.603, align: 'center' });
    cards.forEach(c => s.addText(c.n, Object.assign({ x: c.x + 0.309, y: 4.100, w: 1.662, h: 0.707, fontSize: 36, bold: true }, ctr)));
    navBar(s);
    infographicTitle(s);
}

function slide11(pres) {
    const s = pres.addSlide();
    const panels = [
        { x: 0, fill: BROWN, pct: '80%', barW: 2.571 },
        { x: 6.062, fill: TAN, pct: '64%', barW: 2.004 },
    ];
    panels.forEach(p => {
        s.addShape('round1Rect', { x: 1.264 + p.x, y: 2.648, w: 4.738, h: 1.686, rectRadius: sixth(4.738, 1.686), fill: { color: p.fill }, line: { type: 'none' } });
        s.addText(p.pct, { x: 4.952 + p.x, y: 3.638, w: 0.821, h: 0.352, fontFace: SANS, fontSize: 11, color: WHITE, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
        rounded(s, { x: 1.713 + p.x, y: 3.776, w: 3.151, h: 0.103, r: PILL, fill: WHITE, transparency: 80 });
        rounded(s, { x: 1.712 + p.x, y: 3.776, w: p.barW, h: 0.103, r: PILL, fill: WHITE });
        s.addText('Lorem ipsum dolor sit amet, consectetuer dolor sit ',
            { x: 1.600 + p.x, y: 3.229, w: 3.992, h: 0.627, fontFace: SANS, fontSize: 11, color: WHITE, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
        s.addText('Title Here', { x: 1.600 + p.x, y: 2.910, w: 2.556, h: 0.337, fontFace: SANS, fontSize: 14, bold: true, color: WHITE, valign: 'top' });
    });
    const LOREM = 'Lorem ipsum dolor sit amet, consect dolor sit amet, consect adipiscing elit. dolor dolor sit '
        + 'amet, consect dolor sit dolor sit amet, consect dolor sit amet, consect adipiscing elit. dolor dolor '
        + 'sit amet, dolor sit amet, consect dolor sit amet, ';
    body(s, LOREM, { x: 1.166, y: 4.895, w: 4.889, h: 1.133 });
    body(s, LOREM, { x: 7.275, y: 4.895, w: 4.889, h: 1.133 });
    navBar(s);
    infographicTitle(s);
}

/** One wing of the slide-12 pinwheel: a slanted bar with two rounded ends. */
function blade(slide, o) {
    const w = o.w, h = o.h;
    slide.addShape('custGeom', {
        x: o.x, y: o.y, w: w, h: h, rotate: o.rotate, flipH: o.flipH, flipV: o.flipV,
        fill: { color: o.fill }, line: { type: 'none' },
        points: [
            { x: 0.178 * w, y: 0, moveTo: true },
            { x: 0.910 * w, y: 0 },
            { x: w, y: 0.155 * h, curve: { type: 'cubic', x1: 0.965 * w, y1: 0, x2: w, y2: 0.055 * h } },
            { x: 0.823 * w, y: h },
            { x: 0.092 * w, y: h },
            { x: 0, y: 0.845 * h, curve: { type: 'cubic', x1: 0.035 * w, y1: h, x2: 0, y2: 0.945 * h } },
            { close: true },
        ],
    });
}

function slide12(pres) {
    const s = pres.addSlide();
    blade(s, { x: 2.138, y: 4.695, w: 1.606, h: 1.121, rotate: 270, flipH: true, flipV: true, fill: 'D7CBBC' });
    blade(s, { x: 1.247, y: 3.829, w: 1.663, h: 1.121, rotate: 180, fill: BROWN });
    blade(s, { x: 2.974, y: 3.829, w: 1.660, h: 1.121, rotate: 180, flipH: true, flipV: true, fill: BROWN });
    blade(s, { x: 2.136, y: 2.965, w: 1.610, h: 1.121, rotate: 270, fill: 'B3A093' });
    // Line-chart glyph (bottom wing).
    s.addShape('custGeom', {
        x: 2.723, y: 5.157, w: 0.436, h: 0.372, fill: { color: WHITE }, line: { type: 'none' },
        points: [{ x: 0.00, y: 0.00, moveTo: true }, { x: 0.042, y: 0.00 }, { x: 0.042, y: 0.330 },
            { x: 0.436, y: 0.330 }, { x: 0.436, y: 0.372 }, { x: 0.00, y: 0.372 }, { close: true }],
    });
    s.addShape('custGeom', {
        x: 2.775, y: 5.220, w: 0.340, h: 0.180, fill: { color: WHITE }, line: { type: 'none' },
        points: [{ x: 0.00, y: 0.130, moveTo: true }, { x: 0.120, y: 0.020 }, { x: 0.200, y: 0.090 },
            { x: 0.340, y: 0.00 }, { x: 0.340, y: 0.045 }, { x: 0.205, y: 0.140 }, { x: 0.122, y: 0.070 },
            { x: 0.022, y: 0.180 }, { close: true }],
    });
    // Bar-chart glyph (right wing).
    s.addShape('rect', { x: 3.716, y: 4.203, w: 0.372, h: 0.372, fill: { type: 'none' }, line: { color: WHITE, width: 2 } });
    [[0.055, 0.150], [0.145, 0.230], [0.235, 0.110]].forEach(b => s.addShape('rect', {
        x: 3.716 + b[0], y: 4.203 + 0.310 - b[1], w: 0.062, h: b[1], fill: { color: WHITE }, line: { type: 'none' },
    }));
    // Presenter + speech-bubble glyph (top wing).
    s.addShape('ellipse', { x: 2.990, y: 3.270, w: 0.116, h: 0.118, fill: { color: WHITE }, line: { type: 'none' } });
    rounded(s, { x: 2.947, y: 3.412, w: 0.201, h: 0.362, r: [0.075, 0.075, 0, 0], fill: WHITE });
    s.addShape('rect', { x: 3.041, y: 3.430, w: 0.014, h: 0.130, fill: { color: 'B3A093' }, line: { type: 'none' } });
    s.addShape('custGeom', { // speech bubble with a tail pointing right
        x: 2.678, y: 3.290, w: 0.254, h: 0.176, fill: { color: WHITE }, line: { type: 'none' },
        points: [{ x: 0.030, y: 0.000, moveTo: true }, { x: 0.224, y: 0.000 },
            { x: 0.254, y: 0.030, curve: { type: 'cubic', x1: 0.241, y1: 0.000, x2: 0.254, y2: 0.013 } },
            { x: 0.254, y: 0.176 }, { x: 0.180, y: 0.130 }, { x: 0.030, y: 0.130 },
            { x: 0.000, y: 0.100, curve: { type: 'cubic', x1: 0.013, y1: 0.130, x2: 0.000, y2: 0.117 } },
            { x: 0.000, y: 0.030 },
            { x: 0.030, y: 0.000, curve: { type: 'cubic', x1: 0.000, y1: 0.013, x2: 0.013, y2: 0.000 } },
            { close: true }],
    });
    // Award-rosette glyph (left wing): starburst medal over two ribbon tails.
    s.addShape('triangle', { x: 1.848, y: 4.360, w: 0.110, h: 0.200, fill: { color: WHITE }, line: { type: 'none' }, rotate: 180 });
    s.addShape('triangle', { x: 2.010, y: 4.360, w: 0.110, h: 0.200, fill: { color: WHITE }, line: { type: 'none' }, rotate: 180 });
    s.addShape('star16', { x: 1.835, y: 4.166, w: 0.320, h: 0.320, fill: { color: WHITE }, line: { type: 'none' } });
    s.addShape('ellipse', { x: 1.879, y: 4.210, w: 0.232, h: 0.232, fill: { color: BROWN }, line: { type: 'none' } });
    s.addShape('custGeom', { // thumbs-up inside the medal
        x: 1.925, y: 4.252, w: 0.140, h: 0.148, fill: { color: WHITE }, line: { type: 'none' },
        points: [{ x: 0.000, y: 0.062, moveTo: true }, { x: 0.040, y: 0.062 }, { x: 0.040, y: 0.148 },
            { x: 0.000, y: 0.148 }, { close: true },
            { x: 0.055, y: 0.055, moveTo: true }, { x: 0.085, y: 0.000 }, { x: 0.105, y: 0.010 },
            { x: 0.095, y: 0.055 }, { x: 0.140, y: 0.055 }, { x: 0.125, y: 0.148 }, { x: 0.055, y: 0.148 }, { close: true }],
    });
    // Four labelled paragraphs with a coloured tick bar.
    const cells = [
        { x: 5.968, tx: 6.424, y: 2.666, bar: 'B3A093' },
        { x: 9.544, tx: 10.000, y: 2.666, bar: 'D7CBBC' },
        { x: 5.968, tx: 6.424, y: 4.846, bar: BROWN },
        { x: 9.544, tx: 10.000, y: 4.846, bar: TAN },
    ];
    cells.forEach(c => {
        s.addShape('rect', { x: c.x, y: c.y, w: 0.050, h: 1.012, fill: { color: c.bar }, line: { type: 'none' } });
        s.addText('Title Here', { x: c.tx, y: c.y, w: 1.653, h: 0.303, fontFace: SANS, fontSize: 12, bold: true, color: c.bar, valign: 'top' });
        body(s, 'Lorem ipsum dolor sit amet, consectetuer adi piscing elit. ipsum adi piscing elit. ',
            { x: c.tx, y: c.y + 0.358, w: 2.3334, h: 1.182, fontSize: 11 });
    });
    navBar(s);
    infographicTitle(s);
}

function slide13(pres) {
    const s = pres.addSlide();
    // Per card: the big plate, the flap peeking out below it and the darker
    // wedge that fakes a drop shadow between the two.
    const cards = [
        { x: 1.175, main: 'CDBFB7', flap: 'B3A093', wedge: '5A4A3F', text: BODY, n: '01' },
        { x: 4.011, main: BROWN, flap: '5A4A3F', wedge: '3C312A', text: WHITE, n: '02' },
        { x: 6.847, main: 'D7CBBC', flap: '5A4A3F', wedge: '9B805D', text: WHITE, n: '03' },
        { x: 9.683, main: TAN, flap: '3C312A', wedge: '68553E', text: WHITE, n: '04' },
    ];
    cards.forEach(c => {
        const fx = c.x + 0.317; // flap/wedge sit slightly right of the plate
        rounded(s, { x: fx, y: 4.767, w: 2.159, h: 0.889, r: [0.148, 0, 0, 0.148], fill: c.flap });
        s.addShape('custGeom', {
            x: fx, y: 4.611, w: 2.159, h: 0.889, fill: { color: c.wedge }, line: { type: 'none' },
            points: [{ x: 0, y: 0.117, moveTo: true }, { x: 2.159, y: 0 },
                { x: 2.159, y: 0.889 }, { x: 0, y: 0.772 }, { close: true }],
        });
        s.addShape('round2DiagRect', { x: c.x, y: 2.608, w: 2.476, h: 2.601, fill: { color: c.main }, line: { type: 'none' } });
    });
    cards.forEach(c => {
        s.addText('Step ' + c.n, { x: c.x + 0.276, y: 3.164, w: 1.595, h: 0.438, fontFace: SANS, fontSize: 20, bold: true, color: c.text, align: 'justify', valign: 'top' });
        body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Mae',
            { x: c.x + 0.276, y: 3.625, w: 1.924, h: 0.904, fontSize: 11, color: c.text });
    });
    navBar(s);
    infographicTitle(s);
}

function slide14(pres) {
    const s = pres.addSlide();
    const rows = [
        { y: 2.579, py: 2.569, iy: 2.815, cy: 3.299, tint: BROWN },
        { y: 4.464, py: 4.455, iy: 4.700, cy: 5.185, tint: TAN },
    ];
    rows.forEach(r => s.addShape('roundRect', {
        x: 2.135, y: r.y, w: 1.540, h: 1.647, rectRadius: sixth(1.540, 1.647), rotate: 180, flipH: true,
        fill: { color: r.tint }, line: { type: 'none' }, shadow: shadow(50, 20, 0.2),
    }));
    rows.forEach(r => {
        s.addShape('foldedCorner', {
            x: 4.040, y: r.py, w: 7.159, h: 1.656, rectRadius: 0.828,
            fill: { color: r.tint }, line: { type: 'none' }, shadow: shadow(50, 20, 0.2),
        });
        s.addShape('roundRect', {
            x: 2.363, y: r.iy, w: 1.083, h: 1.165, rectRadius: sixth(1.083, 1.165), rotate: 180,
            fill: { color: 'F2F2F2' }, line: { type: 'none' }, shadow: shadow(50, 20, 0.2),
        });
        // Check mark inside the white tile.
        s.addShape('custGeom', {
            x: 2.612, y: r.cy, w: 0.586, h: 0.397, fill: { color: r.tint }, line: { type: 'none' },
            points: [{ x: 0.574, y: 0.063, moveTo: true }, { x: 0.252, y: 0.385 }, { x: 0.229, y: 0.397 },
                { x: 0.199, y: 0.385 }, { x: 0.021, y: 0.209 }, { x: 0.000, y: 0.176 }, { x: 0.042, y: 0.135 },
                { x: 0.229, y: 0.301 }, { x: 0.523, y: 0.010 }, { x: 0.545, y: 0.000 }, { x: 0.586, y: 0.042 }, { close: true }],
        });
        s.addText('Your Title Here', { x: 4.492, y: r.py + 0.198, w: 3.521, h: 0.370, fontFace: SANS, fontSize: 16, bold: true, color: WHITE, align: 'justify', valign: 'top' });
        body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, '
            + 'magna sedipsum dolor sit a dolor sit amet, consectetuer dolor sit amet, consectetuer ',
            { x: 4.492, y: r.py + 0.555, w: 5.772, h: 0.904, fontSize: 11, color: WHITE });
    });
    navBar(s);
    infographicTitle(s);
}

function slide15(pres) {
    const s = pres.addSlide();
    // The long U-turn ribbon with the arrow head on the upper right.
    const AW = 12.189, AH = 2.167;
    s.addShape('custGeom', {
        x: 1.144, y: 3.108, w: AW, h: AH, fill: { color: BROWN }, line: { type: 'none' },
        points: [
            { x: 0.8699 * AW, y: 0, moveTo: true }, { x: 0.9077 * AW, y: 0.25 * AH }, { x: 0.8699 * AW, y: 0.50 * AH },
            { x: 0.8699 * AW, y: 0.3352 * AH }, { x: 0.0742 * AW, y: 0.3352 * AH },
            { x: 0.0303 * AW, y: 0.5824 * AH, curve: { type: 'cubic', x1: 0.0500 * AW, y1: 0.3352 * AH, x2: 0.0303 * AW, y2: 0.4459 * AH } },
            { x: 0.0742 * AW, y: 0.8296 * AH, curve: { type: 'cubic', x1: 0.0303 * AW, y1: 0.7189 * AH, x2: 0.0500 * AW, y2: 0.8296 * AH } },
            { x: AW, y: 0.8296 * AH }, { x: AW, y: AH }, { x: 0.0742 * AW, y: AH },
            { x: 0, y: 0.5824 * AH, curve: { type: 'cubic', x1: 0.0332 * AW, y1: AH, x2: 0, y2: 0.8130 * AH } },
            { x: 0.0742 * AW, y: 0.1648 * AH, curve: { type: 'cubic', x1: 0, y1: 0.3518 * AH, x2: 0.0332 * AW, y2: 0.1648 * AH } },
            { x: 0.8699 * AW, y: 0.1648 * AH }, { x: 0.8699 * AW, y: 0 }, { close: true },
        ],
    });
    const steps = [
        { cx: 2.089, cy: 2.873, tint: BROWN, tx: 3.159, ty: 2.582, amount: '$88.500.000' },
        { cx: 6.539, cy: 2.873, tint: TAN, tx: 7.609, ty: 2.582, amount: '$89.000.000' },
        { cx: 4.541, cy: 4.675, tint: TAN, tx: 5.612, ty: 5.348, amount: '$90.500.000' },
        { cx: 8.992, cy: 4.675, tint: BROWN, tx: 10.062, ty: 5.348, amount: '$90.000.000' },
    ];
    steps.forEach(st => {
        s.addShape('chevron', { x: st.cx, y: st.cy, w: 1.069, h: 1.180, fill: { color: st.tint }, line: { type: 'none' }, shadow: shadow(9, 5, 0.2) });
        s.addShape('ellipse', { x: st.cx - 0.273, y: st.cy + 0.114, w: 0.945, h: 0.945, fill: { color: WHITE }, line: { type: 'none' }, shadow: shadow(9, 5, 0.2) });
        s.addShape('ellipse', { x: st.cx - 0.243, y: st.cy + 0.145, w: 0.885, h: 0.885, fill: { color: st.tint }, line: { type: 'none' } });
        // Miniature bar chart inside the circle.
        const ix = st.cx + 0.055, iy = st.cy + 0.466;
        s.addShape('rect', { x: ix, y: iy + 0.221, w: 0.294, h: 0.028, fill: { color: WHITE }, line: { type: 'none' } });
        [[0.041, 0.076], [0.121, 0.141], [0.197, 0.207]].forEach(b => s.addShape('rect', {
            x: ix + b[0], y: iy + 0.207 - b[1], w: 0.055, h: b[1], fill: { color: WHITE }, line: { type: 'none' },
        }));
        s.addText(st.amount, { x: st.tx, y: st.ty, w: 1.498, h: 0.303, fontFace: SANS, fontSize: 12, bold: true, color: st.tint, align: 'justify', valign: 'top' });
        body(s, 'Lorem ipsum dolor sit amet, consect dolor sit.', { x: st.tx, y: st.ty + 0.242, w: 2.198, h: 0.603 });
    });
    navBar(s);
    infographicTitle(s);
}

function slide16(pres) {
    const s = pres.addSlide();
    const groups = [
        { x: 0, outer: BROWN, pct: '47%', tint: BROWN },
        { x: 3.680, outer: TAN, pct: '83%', tint: TAN },
        { x: 7.360, outer: BROWN, pct: '19%', tint: BROWN },
    ];
    groups.forEach(g => {
        const dy = g.x === 0 ? 0 : (g.x === 3.680 ? 0.003 : 0.006);
        rounded(s, { x: 1.917 + g.x, y: 2.764 + dy, w: 2.139, h: 2.139, r: [0, sixth(2.139, 2.139), 0, sixth(2.139, 2.139)], fill: g.outer });
        // Six little capsules poking out of the badge, three per side.
        const caps = [
            { x: 3.746, y: 3.081, w: 1.039, c: 'CDBFB7', f: false },
            { x: 3.746, y: 3.545, w: 0.726, c: g.outer, f: false },
            { x: 3.746, y: 4.008, w: 1.039, c: 'CDBFB7', f: false },
            { x: 1.501, y: 3.087, w: 0.726, c: 'CDBFB7', f: true },
            { x: 1.188, y: 3.551, w: 1.039, c: g.outer, f: true },
            { x: 1.501, y: 4.014, w: 0.726, c: 'CDBFB7', f: true },
        ];
        caps.forEach(c => s.addShape('flowChartTerminator', {
            x: c.x + g.x, y: c.y + dy, w: c.w, h: 0.325, flipH: c.f,
            fill: { color: c.c }, line: { type: 'none' },
        }));
        rounded(s, {
            x: 2.070 + g.x, y: 2.917 + dy, w: 1.833, h: 1.833, fill: BROWN,
            r: [0, sixth(1.833, 1.833), 0, sixth(1.833, 1.833)],
            shadow: shadow(13, 3, 0.2),
        });
        s.addText(g.pct, { x: 2.427 + g.x, y: 3.229 + dy, w: 1.185, h: 0.640, fontFace: SANS, fontSize: 32, bold: true, color: WHITE, valign: 'top' });
        s.addText('Lorem ipsum dolor sit ipsum dolor', {
            x: 2.223 + g.x, y: 3.835 + dy, w: 1.527, h: 0.867, fontFace: SANS, fontSize: 10.5,
            color: WHITE, align: 'center', lineSpacingMultiple: 1.5, valign: 'top',
        });
        s.addText([
            { text: g.pct + ' ', options: { bold: true } },
            { text: 'In Progress', options: { bold: false } },
        ], { x: 1.859 + g.x, y: 5.196 + dy, w: 2.255, h: 0.337, fontFace: SANS, fontSize: 14, color: g.tint, align: 'center', valign: 'top' });
        s.addText('Lorem ipsum dolor sit amet, consectetuer adipiscing. elit. Maecenas porttitor.', {
            x: 1.313 + g.x, y: 5.516 + dy, w: 3.347, h: 0.603, fontFace: SANS, fontSize: 10.5,
            color: BODY, align: 'center', lineSpacingMultiple: 1.5, valign: 'top',
        });
    });
    navBar(s);
    infographicTitle(s);
}

function slide17(pres) {
    const s = pres.addSlide();
    // Right-hand legend: six lettered bullets in two columns.
    const legend = [
        { x: 7.634, y: 2.896, tag: 'A', tint: BROWN }, { x: 10.697, y: 2.896, tag: 'D', tint: TAN },
        { x: 7.634, y: 4.172, tag: 'B', tint: TAN }, { x: 10.697, y: 4.172, tag: 'E', tint: BROWN },
        { x: 7.634, y: 5.448, tag: 'C', tint: BROWN }, { x: 10.697, y: 5.448, tag: 'F', tint: TAN },
    ];
    legend.forEach(it => {
        body(s, 'Lorem ipsum sit amet consect ', { x: it.x, y: it.y, w: 1.529, h: 0.627, fontSize: 11 });
        s.addShape('ellipse', { x: it.x - 0.322, y: it.y + 0.107, w: 0.302, h: 0.302, fill: { color: it.tint }, line: { type: 'none' } });
        s.addText(it.tag, { x: it.x - 0.322, y: it.y + 0.107, w: 0.302, h: 0.302, align: 'center', valign: 'middle', fontFace: SANS, fontSize: 10, color: WHITE });
    });
    // Two percentage cards.
    const cards = [
        { x: 1.195, tint: BROWN, pct: '67%', tag: '1', flip: false },
        { x: 4.061, tint: TAN, pct: '93%', tag: '2', flip: true },
    ];
    cards.forEach(c => {
        s.addShape('round1Rect', {
            x: c.x, y: 2.979, w: 2.396, h: 2.161, rectRadius: sixth(2.396, 2.161), flipH: c.flip,
            fill: { color: c.tint }, line: { type: 'none' },
        });
        s.addText(c.pct, { x: c.x + 0.232, y: 3.512, w: 1.932, h: 0.774, align: 'center', fontFace: SANS, fontSize: 40, bold: true, color: WHITE, valign: 'top' });
        s.addText('Title Here', { x: c.x + 0.506, y: 4.193, w: 1.384, h: 0.370, align: 'center', fontFace: SANS, fontSize: 16, italic: true, color: WHITE, valign: 'top' });
        body(s, 'Lorem ipsum dolor sit amet, consectetuer', { x: c.x + 0.327, y: 5.492, w: 2.069, h: 0.627, fontSize: 11 });
        s.addShape('ellipse', { x: c.x + 0.005, y: 5.599, w: 0.302, h: 0.302, fill: { color: c.tint }, line: { type: 'none' } });
        s.addText(c.tag, { x: c.x + 0.005, y: 5.599, w: 0.302, h: 0.302, align: 'center', valign: 'middle', fontFace: SANS, fontSize: 10, color: WHITE });
    });
    navBar(s);
    infographicTitle(s);
}

function slide18(pres) {
    const s = pres.addSlide();
    // Two long bars crossed at 45 degrees form the "X".
    s.addShape('rect', { x: 6.137, y: 2.545, w: 1.060, h: 4.499, rotate: 45, fill: { color: BROWN }, line: { type: 'none' } });
    s.addShape('rect', { x: 6.128, y: 2.580, w: 1.077, h: 4.429, rotate: 135, fill: { color: TAN }, line: { type: 'none' } });
    const nodes = [
        { x: 4.528, y: 2.651, icon: 'folder', tint: TAN },
        { x: 7.728, y: 2.651, icon: 'gears', tint: BROWN },
        { x: 4.528, y: 5.861, icon: 'org', tint: BROWN },
        { x: 7.728, y: 5.861, icon: 'ball', tint: TAN },
    ];
    nodes.forEach(n => {
        s.addShape('ellipse', { x: n.x, y: n.y, w: 1.077, h: 1.077, fill: { color: WHITE }, line: { type: 'none' }, shadow: shadow(12, 0, 0.33) });
        s.addShape('donut', { x: n.x, y: n.y, w: 1.077, h: 1.077, fill: { color: WHITE }, line: { type: 'none' } });
        const cx = n.x + 0.333, cy = n.y + 0.334;
        if (n.icon === 'folder') {
            rounded(s, { x: cx, y: cy + 0.078, w: 0.411, h: 0.233, r: [0.030, 0.030, 0.030, 0.030], fill: n.tint });
            s.addShape('rect', { x: cx, y: cy + 0.020, w: 0.175, h: 0.075, fill: { color: n.tint }, line: { type: 'none' } });
        } else if (n.icon === 'gears') {
            s.addShape('gear9', { x: cx, y: cy + 0.045, w: 0.284, h: 0.284, fill: { color: n.tint }, line: { type: 'none' } });
            s.addShape('gear9', { x: cx + 0.231, y: cy - 0.085, w: 0.181, h: 0.181, fill: { color: n.tint }, line: { type: 'none' } });
        } else if (n.icon === 'org') {
            s.addShape('rect', { x: cx + 0.120, y: cy - 0.010, w: 0.170, h: 0.105, fill: { color: n.tint }, line: { type: 'none' } });
            s.addShape('rect', { x: cx, y: cy + 0.235, w: 0.150, h: 0.105, fill: { color: n.tint }, line: { type: 'none' } });
            s.addShape('rect', { x: cx + 0.255, y: cy + 0.235, w: 0.150, h: 0.105, fill: { color: n.tint }, line: { type: 'none' } });
            s.addShape('rect', { x: cx + 0.072, y: cy + 0.140, w: 0.262, h: 0.020, fill: { color: n.tint }, line: { type: 'none' } });
            [0.072, 0.196, 0.324].forEach(dx => s.addShape('rect', {
                x: cx + dx, y: cy + 0.095, w: 0.018, h: 0.140, fill: { color: n.tint }, line: { type: 'none' },
            }));
        } else {
            // Dribbble-style ball: outline circle crossed by three arcs.
            s.addShape('ellipse', { x: cx + 0.020, y: cy + 0.020, w: 0.370, h: 0.370, fill: { type: 'none' }, line: { color: n.tint, width: 2 } });
            s.addShape('arc', { x: cx - 0.150, y: cy - 0.020, w: 0.420, h: 0.420, angleRange: [300, 60], line: { color: n.tint, width: 1.5 } });
            s.addShape('arc', { x: cx + 0.090, y: cy + 0.150, w: 0.420, h: 0.420, angleRange: [130, 250], line: { color: n.tint, width: 1.5 } });
            s.addShape('arc', { x: cx - 0.020, y: cy - 0.180, w: 0.460, h: 0.460, angleRange: [40, 150], line: { color: n.tint, width: 1.5 } });
        }
    });
    s.addShape('roundRect', {
        x: 6.042, y: 4.170, w: 1.249, h: 1.249, rectRadius: sixth(1.249, 1.249), rotate: 45,
        fill: { color: WHITE }, line: { type: 'none' },
        shadow: shadow(30, 0, 0.33),
    });
    s.addText('Your Title', { x: 6.160, y: 4.408, w: 1.014, h: 0.774, align: 'center', fontFace: SANS, fontSize: 20, bold: true, color: BROWN, valign: 'top' });
    const quadrants = [
        { nx: 1.144, tx: 1.848, bx: 1.171, y: 2.654, num: '01', name: 'Strength', tint: TAN, nw: 0.758, tw: 1.151 },
        { nx: 1.144, tx: 1.848, bx: 1.171, y: 4.702, num: '03', name: 'Threats', tint: BROWN, nw: 0.758, tw: 1.151 },
        { nx: 9.695, tx: 10.400, bx: 9.723, y: 2.654, num: '02', name: 'Weaknesses', tint: BROWN, nw: 0.758, tw: 1.526 },
        { nx: 9.695, tx: 10.400, bx: 9.723, y: 4.702, num: '04', name: 'Opportunities', tint: TAN, nw: 0.878, tw: 1.663 },
    ];
    quadrants.forEach(q => {
        s.addText(q.num, { x: q.nx, y: q.y, w: q.nw, h: 0.640, align: 'center', fontFace: SANS, fontSize: 32, bold: true, color: q.tint, valign: 'top' });
        s.addText(q.name, { x: q.tx, y: q.y + 0.152, w: q.tw, h: 0.337, fontFace: SANS, fontSize: 14, bold: true, color: q.tint, valign: 'top' });
        body(s, 'Lorem ipsum dolor sit amet, consectetuer adipi scing elit. Lorem ipsum dolor sit ',
            { x: q.bx, y: q.y + 0.584, w: 2.467, h: 0.904, fontSize: 11 });
    });
    navBar(s);
    infographicTitle(s);
}

/** Phone mock-up: near-black body, grey screen, notch across the top. */
function phone(slide, x, y, w, h) {
    const bez = 0.075, r = w * 0.155;
    rounded(slide, { x: x, y: y, w: w, h: h, r: [r, r, r, r], fill: '17181A' });
    photo(slide, {
        x: x + bez, y: y + bez, w: w - 2 * bez, h: h - 2 * bez,
        r: [r - bez, r - bez, r - bez, r - bez],
    });
    rounded(slide, {
        x: x + w * 0.29, y: y + bez, w: w * 0.42, h: 0.14, fill: '17181A',
        r: [0, 0, 0.07, 0.07],
    });
}

function slide19(pres) {
    const s = pres.addSlide();
    phone(s, 9.540, 1.348, 2.620, 4.744);
    navBar(s);
    heading(s, 'Contact ', 'Us', { x: 1.045, y: 1.222, w: 4.662, h: 0.774 });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ', { x: 1.055, y: 2.266, w: 4.050, h: 0.601 });
    s.addShape('round1Rect', { x: 0, y: 3.75, w: 5.105, h: 3.75, rectRadius: 1.875, fill: { color: BROWN }, line: { type: 'none' } });
    const lines = [
        { y: 4.435, t: '+12345678910', w: 1.633, icon: 'phone' },
        { y: 4.966, t: 'yourmail@gmail.com', w: 1.953, icon: 'mail' },
        { y: 5.552, t: 'www.yourwebsite.com', w: 1.953, icon: 'web' },
        { y: 6.139, t: '123 Anywhere ST,, Any City', w: 2.259, icon: 'home' },
    ];
    lines.forEach(l => {
        s.addText(l.t, { x: 1.653, y: l.y, w: l.w, h: 0.278, fontFace: SANS, fontSize: 10.5, color: WHITE, align: 'justify', valign: 'top' });
        const ix = 1.162, iy = l.y + 0.039;
        if (l.icon === 'phone') {
            s.addShape('custGeom', {
                x: ix, y: iy, w: 0.222, h: 0.200, fill: { color: WHITE }, line: { type: 'none' },
                points: [{ x: 0.158, y: 0.110, moveTo: true }, { x: 0.135, y: 0.131 }, { x: 0.095, y: 0.110 },
                    { x: 0.071, y: 0.069 }, { x: 0.095, y: 0.048 }, { x: 0.032, y: 0.000 }, { x: 0.008, y: 0.021 },
                    { x: 0.000, y: 0.062 }, { x: 0.063, y: 0.131 }, { x: 0.198, y: 0.200 }, { x: 0.222, y: 0.159 }, { close: true }],
            });
        } else if (l.icon === 'mail') {
            s.addShape('rect', { x: ix, y: iy, w: 0.210, h: 0.132, fill: { type: 'none' }, line: { color: WHITE, width: 1.25 } });
            s.addShape('line', { x: ix, y: iy, w: 0.105, h: 0.066, line: { color: WHITE, width: 1.25 } });
            s.addShape('line', { x: ix + 0.105, y: iy, w: 0.105, h: 0.066, line: { color: WHITE, width: 1.25 }, flipH: true });
        } else if (l.icon === 'web') {
            s.addShape('ellipse', { x: ix, y: iy, w: 0.214, h: 0.185, fill: { type: 'none' }, line: { color: WHITE, width: 1.25 } });
            s.addShape('ellipse', { x: ix + 0.062, y: iy, w: 0.090, h: 0.185, fill: { type: 'none' }, line: { color: WHITE, width: 1 } });
            s.addShape('line', { x: ix, y: iy + 0.092, w: 0.214, h: 0, line: { color: WHITE, width: 1 } });
        } else {
            s.addShape('triangle', { x: ix, y: iy, w: 0.223, h: 0.095, fill: { color: WHITE }, line: { type: 'none' } });
            s.addShape('rect', { x: ix + 0.045, y: iy + 0.090, w: 0.133, h: 0.075, fill: { color: WHITE }, line: { type: 'none' } });
        }
    });
    dots(s, 5.707, 3.750, 'v');
    rounded(s, { x: 7.451, y: 6.821, w: 2.535, h: 0.100, r: PILL, fill: TAN });
    rounded(s, { x: 9.809, y: 6.822, w: 2.367, h: 0.100, r: PILL, fill: BROWN });
    phone(s, 8.227, 2.193, 1.733, 3.520);
}

function slide20(pres) {
    const s = pres.addSlide();
    rounded(s, { x: 1.158, y: 1.411, w: 4.200, h: 4.705, r: ARCH, fill: BROWN });
    photo(s, { x: 1.241, y: 1.481, w: 4.034, h: 4.539, r: ARCH });
    navBar(s);
    rounded(s, { x: 3.541, y: 2.605, w: 3.919, h: 2.316, r: PILL, fill: WHITE });
    s.addText([
        { text: 'Thank ', options: { color: TAN } },
        { text: 'You', options: { color: BROWN } },
    ], { x: 4.010, y: 3.157, w: 6.643, h: 1.313, fontFace: HEAD, fontSize: 72, bold: true, valign: 'top' });
    rounded(s, { x: 7.451, y: 6.061, w: 2.535, h: 0.100, r: PILL, fill: TAN });
    rounded(s, { x: 9.809, y: 6.062, w: 2.367, h: 0.100, r: PILL, fill: BROWN });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ', { x: 8.147, y: 1.244, w: 4.050, h: 0.601 });
}

/* ------------------------------------------------------------------ build */

function build() {
    const pres = new PptxGenJS();
    pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5 in
    pres.author = 'pptxgenjs';
    pres.title = 'Aesthetic - Photography Presentation Template';

    [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
        slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
        .forEach(fn => fn(pres));

    return pres.writeFile({
        fileName: path.join(__dirname, '113d14f4-c35c-44b5-8bf4-4bdc9e1b74aa_grok_final.pptx'),
    });
}

build().then(f => console.log('wrote ' + f)).catch(e => { console.error(e); process.exit(1); });
