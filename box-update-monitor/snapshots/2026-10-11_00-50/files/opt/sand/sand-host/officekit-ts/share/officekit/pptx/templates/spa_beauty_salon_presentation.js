/**
 * "Melati - Spa & Beauty Presentation" (15 slides, 13.333in x 7.5in)
 * Rebuilt with pptxgenjs. Raster photos in the original deck are replaced by
 * flat grey placeholder shapes that keep the original position / silhouette.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const PINK = 'F37AAB';   // theme accent1
const PALE_PINK = 'FAC8DC';   // footer url
const INK = '3F3F3F';   // body dark grey
const INK_BLACK = '262626';   // heading near-black
const MUTED = 'A5A5A5';   // paragraph grey
const BG = 'F2F2F2';   // slide background
const WHITE = 'FFFFFF';
const PHOTO = 'CCCCCC';   // image placeholder fill
const PHOTO_TXT = 'B3B3B3';   // "[image]" caption
const DEVICE_BLACK = '1A1A1A';   // device bezel in mockups
const DEVICE_SILVER = 'DDE1E0';   // device back shell

const SERIF = 'Playfair Display';
const DISPLAY = 'DM Serif Display';
const SANS = 'Poppins';
const SANS_SB = 'Poppins SemiBold';

// Google-Slides text insets (0.1in sides / 0.05in top-bottom) expressed in points
const INSET = [7.2, 7.2, 3.6, 3.6];

const TITLE_SIZE = 44;   // every section heading
const BODY_SIZE = 12;
const LEAD = 1.5;  // 150% line spacing used by all body copy

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

/** Blend two hex colours; t = 0 gives `a`, t = 1 gives `b`. */
function mix(a, b, t) {
    let out = '';
    for (let i = 0; i < 6; i += 2) {
        const va = parseInt(a.substr(i, 2), 16), vb = parseInt(b.substr(i, 2), 16);
        out += Math.round(va + (vb - va) * t).toString(16).toUpperCase().padStart(2, '0');
    }
    return out;
}

/** Text box that behaves like the original (top anchored, Google insets). */
function text(slide, content, opts) {
    slide.addText(content, Object.assign({
        valign: 'top', margin: INSET, fontFace: SANS, fontSize: BODY_SIZE, color: INK,
    }, opts));
}

/** Section heading: array of {text, color?, breakLine?} runs at 44pt Playfair. */
function heading(slide, x, y, w, h, runs, opts) {
    const parts = runs.map(r => ({
        text: r.text,
        options: { color: r.color || INK, breakLine: !!r.breakLine },
    }));
    text(slide, parts, Object.assign({ x, y, w, h, fontFace: SERIF, fontSize: TITLE_SIZE, color: INK }, opts));
}

/** Grey paragraph copy at 12pt / 150% leading. */
function body(slide, x, y, w, h, str, opts) {
    text(slide, str, Object.assign({ x, y, w, h, color: MUTED, lineSpacingMultiple: LEAD }, opts));
}

/** "[image]" caption drawn on top of a placeholder shape. */
function photoLabel(slide, cx, cy) {
    text(slide, '[image]', {
        x: cx - 0.9, y: cy - 0.14, w: 1.8, h: 0.28,
        align: 'center', fontSize: 11, color: PHOTO_TXT, fontFace: SANS,
    });
}

/** Rectangular / rounded image placeholder. */
function photoRect(slide, x, y, w, h, o) {
    const opt = o || {};
    slide.addShape(opt.radius ? 'roundRect' : 'rect', {
        x, y, w, h, fill: { color: opt.color || PHOTO },
        rectRadius: opt.radius || undefined,
    });
    if (opt.label !== false) photoLabel(slide, x + w / 2, y + h / 2);
}

/** Circular / elliptical image placeholder. */
function photoEllipse(slide, x, y, w, h, o) {
    const opt = o || {};
    slide.addShape('ellipse', { x, y, w, h, fill: { color: opt.color || PHOTO } });
    if (opt.label !== false) photoLabel(slide, x + w / 2, y + h / 2);
}

/** Free-form quadrilateral (used by the tilted device mock-ups). */
function quad(slide, pts, fill) {
    const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
    const x = Math.min.apply(null, xs), y = Math.min.apply(null, ys);
    const w = Math.max.apply(null, xs) - x, h = Math.max.apply(null, ys) - y;
    slide.addShape('custGeom', {
        x, y, w, h, fill: { color: fill },
        points: pts.map(p => ({ x: p[0] - x, y: p[1] - y })).concat([{ close: true }]),
    });
}

/** Offset every edge of a convex quad inwards by `d` inches (device bezel). */
function insetQuad(pts, d) {
    const cx = pts.reduce((a, p) => a + p[0], 0) / 4;
    const cy = pts.reduce((a, p) => a + p[1], 0) / 4;
    const lines = pts.map((p, i) => {
        const q = pts[(i + 1) % 4];
        let nx = -(q[1] - p[1]), ny = q[0] - p[0];
        const len = Math.hypot(nx, ny);
        nx /= len; ny /= len;
        if ((cx - p[0]) * nx + (cy - p[1]) * ny < 0) { nx = -nx; ny = -ny; }
        return { nx, ny, c: nx * p[0] + ny * p[1] + d };
    });
    // each output corner is the crossing of the two edges that met there
    return pts.map((p, i) => {
        const a = lines[(i + 3) % 4], b = lines[i];
        const det = a.nx * b.ny - a.ny * b.nx;
        return [(a.c * b.ny - b.c * a.ny) / det, (a.nx * b.c - b.nx * a.c) / det];
    });
}

/**
 * Free-form silhouette in 0..1 shape-relative units.
 * `[x, y]` = straight segment, `[x, y, c1x, c1y, c2x, c2y]` = cubic bezier.
 */
function blob(slide, x, y, w, h, fill, path) {
    slide.addShape('custGeom', {
        x, y, w, h, fill: { color: fill },
        points: path.map((p, i) => (p.length === 6
            ? { x: p[0] * w, y: p[1] * h, curve: { type: 'cubic', x1: p[2] * w, y1: p[3] * h, x2: p[4] * w, y2: p[5] * h } }
            : { x: p[0] * w, y: p[1] * h, moveTo: i === 0 }
        )).concat([{ close: true }]),
    });
}

/** Grey silhouette standing in for a cropped photo, with an "[image]" caption. */
function photoBlob(slide, x, y, w, h, path, labelAt) {
    blob(slide, x, y, w, h, PHOTO, path);
    if (labelAt) photoLabel(slide, x + labelAt[0] * w, y + labelAt[1] * h);
}

/* Silhouettes the original deck cuts its photos into (normalised 0..1). */
const SHAPES = {
    // half-pill: full round on the right, flat on the left edge  (slide 2)
    leftHalfPill: [[0.340, 0], [1, 0.5, 0.704, 0, 1, 0.224], [0.340, 1, 1, 0.776, 0.704, 1],
        [0.025, 0.940, 0.226, 1, 0.118, 0.978], [0, 0.928], [0, 0.072], [0.025, 0.060],
        [0.340, 0, 0.118, 0.022, 0.226, 0]],
    // arch: semicircular top on a rectangle                       (slide 3)
    arch: [[0.5, 0], [1, 0.499, 0.776, 0, 1, 0.224], [1, 1, 1, 0.666, 1, 0.833], [0, 1],
        [0, 0.499], [0.5, 0, 0, 0.224, 0.224, 0]],
    // rectangle whose bottom is a half circle                     (slide 6)
    roundBottom: [[0, 0], [1, 0], [1, 0.618], [0.5, 1, 1, 0.829, 0.776, 1],
        [0, 0.618, 0.224, 1, 0, 0.829], [0, 0, 0, 0.412, 0, 0.206]],
    // big soft blob, flat along the left edge                     (slide 13)
    leftBlob: [[0.227, 0], [1, 0.636, 0.654, 0, 1, 0.285], [0.868, 0.992, 1, 0.768, 0.951, 0.890],
        [0.860, 1], [0, 1], [0, 0.028], [0.034, 0.020], [0.227, 0, 0.096, 0.007, 0.161, 0]],
    // quarter-round panel anchored to the top-right corner        (slide 14)
    topRightPanel: [[0.122, 0], [1, 0], [1, 0.892], [0.6412, 1, 0.897, 0.960, 0.774, 1],
        [0, 0.369, 0.287, 1, 0, 0.718], [0.110, 0.016, 0, 0.238, 0.040, 0.117]],
};

/* ---------------------------- icon vocabulary ---------------------- */

/** Five-petal cherry-blossom mark used as the "Melati" logo. */
function flowerLogo(slide, x, y, size) {
    const cx = x + size / 2, cy = y + size / 2;
    const petalW = size * 0.40, petalH = size * 0.60, orbit = size * 0.21;
    for (let i = 0; i < 5; i++) {
        const a = (i * 72 - 90) * Math.PI / 180;
        slide.addShape('ellipse', {
            x: cx + orbit * Math.cos(a) - petalW / 2,
            y: cy + orbit * Math.sin(a) - petalH / 2,
            w: petalW, h: petalH, rotate: i * 72, fill: { color: PINK },
        });
    }
    slide.addShape('ellipse', {
        x: cx - size * 0.08, y: cy - size * 0.08, w: size * 0.16, h: size * 0.16,
        fill: { color: BG },
    });
}

/** Single tick mark drawn from two strokes. */
function checkMark(slide, x, y, size, color, weight) {
    const w = weight || 2;
    slide.addShape('line', { x: x, y: y + size * 0.45, w: size * 0.33, h: size * 0.35, line: { color, width: w } });
    slide.addShape('line', { x: x + size * 0.33, y: y + size * 0.10, w: size * 0.57, h: size * 0.70, line: { color, width: w }, flipV: true });
}

/** The deck's recurring "double tick" bullet icon. */
function doubleCheck(slide, x, y, size, color) {
    checkMark(slide, x, y + size * 0.10, size * 0.78, color, 2.5);
    checkMark(slide, x + size * 0.30, y, size * 0.78, color, 2.5);
}

/** Pointed leaf: blade in the upper right, short stem running down to the left. */
function leafIcon(slide, x, y, w, h, color) {
    const bw = w * 0.82, bh = h * 0.82;
    slide.addShape('custGeom', {
        x: x + w - bw, y, w: bw, h: bh, fill: { color },
        points: [
            { x: 0, y: bh },
            { x: bw, y: 0, curve: { type: 'cubic', x1: 0, y1: bh * 0.28, x2: bw * 0.52, y2: 0 } },
            { x: 0, y: bh, curve: { type: 'cubic', x1: bw, y1: bh * 0.62, x2: bw * 0.42, y2: bh } },
            { close: true },
        ],
    });
    slide.addShape('line', {
        x, y: y + h * 0.82, w: w * 0.18, h: h * 0.18,
        line: { color, width: 1.25 }, flipV: true,
    });
}

/** Three stacked ripples. */
function wavesIcon(slide, x, y, w, h, color) {
    for (let i = 0; i < 3; i++) {
        const yy = y + i * h * 0.38;
        slide.addShape('custGeom', {
            x, y: yy, w, h: h * 0.26, fill: { color },
            points: [
                { x: 0, y: h * 0.22 },
                { x: w, y: 0, curve: { type: 'cubic', x1: w * 0.3, y1: -h * 0.16, x2: w * 0.7, y2: h * 0.2 } },
                { x: w, y: h * 0.26 },
                { x: 0, y: h * 0.26 + h * 0.22, curve: { type: 'cubic', x1: w * 0.7, y1: h * 0.46, x2: w * 0.3, y2: h * 0.10 } },
                { close: true },
            ],
        });
    }
}

/* ---------------------------- page furniture ----------------------- */

/** Logo + wordmark, top-left (inherited from the slide master). */
function brand(slide) {
    flowerLogo(slide, 0.505, 0.414, 0.252);
    text(slide, 'MELATI', { x: 0.755, y: 0.389, w: 1.146, h: 0.303, fontFace: SANS_SB, color: INK });
}

/** Italic url, bottom-right (inherited from the slide master). */
function websiteFooter(slide, color) {
    text(slide, 'www.yourwebsite.com', {
        x: 9.41, y: 6.833, w: 3.487, h: 0.303,
        align: 'right', italic: true, fontFace: SANS_SB, color: color || PALE_PINK,
    });
}

/** Top navigation used by the opening and closing slides. */
const NAV_ITEMS = [
    { label: 'Home', x: 7.675, active: true },
    { label: 'Service', x: 9.034 },
    { label: 'About', x: 10.393 },
    { label: 'Contact Us', x: 11.752 },
];

function navBar(slide) {
    NAV_ITEMS.forEach(item => {
        text(slide, item.label, {
            x: item.x, y: 0.388, w: 1.146, h: 0.303, align: 'center',
            fontFace: SANS_SB, color: item.active ? PINK : INK,
            underline: item.active ? { style: 'sng' } : undefined,
        });
    });
}

/** Pill button: filled or outlined. */
function button(slide, x, y, w, h, label, style) {
    const filled = style !== 'outline';
    slide.addShape('roundRect', {
        x, y, w, h, rectRadius: 0.0783,
        fill: filled ? { color: PINK } : { type: 'none' },
        line: filled ? { type: 'none' } : { color: PINK, width: 1 },
    });
    text(slide, label, {
        x: x - 0.08, y: y + 0.098, w: w + 0.16, h: 0.303,
        align: 'center', bold: true, fontFace: SANS_SB, color: filled ? WHITE : INK,
    });
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

function slide01(pptx) {                                  // Title / hero
    const s = pptx.addSlide();
    s.background = { color: BG };

    [[6.701, 2.235, 7.155], [7.576, 3.601, 4.746]].forEach(([x, y, d]) => {
        s.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: PINK, width: 1 } });
    });

    brand(s);
    navBar(s);

    text(s, 'Melati', { x: 1.117, y: 1.608, w: 6.208, h: 2.036, fontFace: SERIF, fontSize: 115, bold: true, color: PINK });
    text(s, 'Spa & Beauty Presentation', { x: 1.17, y: 3.416, w: 5.886, h: 0.467, fontFace: SANS_SB, fontSize: 16, color: INK, lineSpacingMultiple: LEAD });
    body(s, 1.17, 4.302, 5.382, 0.678,
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt labore dolore magna.');

    button(s, 1.25, 5.4, 1.828, 0.47, 'More Information');
    button(s, 3.32, 5.4, 1.602, 0.47, 'Start Slide', 'outline');

    s.addShape('ellipse', { x: 7.118, y: 2.858, w: 1.049, h: 1.049, fill: { color: PINK } });
    leafIcon(s, 7.403, 3.143, 0.48, 0.48, BG);
}

function slide02(pptx) {                                  // Welcome To
    const s = pptx.addSlide();
    s.background = { color: BG };

    photoBlob(s, 0, 0.982, 4.218, 5.572, SHAPES.leftHalfPill, [0.35, 0.5]);
    s.addShape('ellipse', { x: 2.573, y: 2.009, w: 3.489, h: 3.489, fill: { color: PINK } });
    photoEllipse(s, 2.716, 2.152, 3.202, 3.202, { label: false });

    heading(s, 6.667, 1.691, 5.798, 1.582, [
        { text: 'Welcome To', breakLine: true },
        { text: 'Melati ' },
        { text: 'Presentation', color: PINK },
    ]);
    body(s, 6.667, 3.75, 5.885, 0.981,
        'Lorem ipsum dolor sit amet, consectetuer adipiscing Maecenas porttitor congue massa. ' +
        'Fusce posuere congue massa. Fusce posuere, magna sed pulvinar osuere, posuere');
    body(s, 6.667, 5.015, 5.885, 0.678,
        'Fusce posuere, magna pulvinaposuere congue. Fusce posuere magna Fusce posuere, magna sed pulvinar osuere, ');

    brand(s);
    websiteFooter(s);
}

function slide03(pptx) {                                  // The Beauty Solution
    const s = pptx.addSlide();
    s.background = { color: BG };

    s.addShape('rect', { x: 9.23, y: 0, w: 4.103, h: 7.5, fill: { color: PINK } });

    photoBlob(s, 6.95, 1.176, 4.729, 5.149, SHAPES.arch, [0.5, 0.5]);

    heading(s, 1.069, 1.829, 5.08, 1.582, [
        { text: 'The Beauty ' },
        { text: 'Soluti0n', color: PINK },
        { text: ' For You' },
    ]);

    [['4.7', 'Client Rate', 1.069, 0.693], ['48K', 'Customers', 3.348, 0.91]].forEach(([num, cap, x, w]) => {
        text(s, num, { x, y: 3.803, w, h: 0.572, fontFace: SERIF, fontSize: 28, bold: true, color: PINK });
        text(s, cap, { x, y: 4.359, w: 1.784, h: 0.421, fontFace: SANS_SB, fontSize: 14, italic: true, color: INK, lineSpacingMultiple: LEAD });
    });

    body(s, 1.065, 4.877, 4.935, 0.981,
        'Lorem ipsum dolor sit amet, consectetuer adipiscing Maecenas porttitor congue massa. ' +
        'Fusce posuere congue massa. Fusce posuere magna.');

    brand(s);
    websiteFooter(s, WHITE);
}

function slide04(pptx) {                                  // Beautyful You Are With Us
    const s = pptx.addSlide();
    s.background = { color: BG };

    photoRect(s, 0.682, 1.286, 5.8, 4.128, { radius: 0.203 });
    s.addShape('roundRect', { x: 5.194, y: 4.236, w: 3.588, h: 2.172, rectRadius: 0.2064, fill: { color: PINK } });

    heading(s, 7.23, 1.462, 5.08, 1.582, [
        { text: 'Beautyful You', breakLine: true },
        { text: 'Are ' },
        { text: 'With Us', color: PINK },
    ]);
    body(s, 7.23, 3.227, 5.437, 0.678,
        'Lorem ipsum dolor sit amet, consectetuer adipiscing Maecenas porttitor congue massa posuere.');

    const REASONS = [
        { title: 'First Reason', tx: 6.095, bx: 5.587, color: WHITE, bodyColor: WHITE },
        { title: 'Second Reason', tx: 9.77, bx: 9.262, color: INK, bodyColor: MUTED },
    ];
    REASONS.forEach(r => {
        text(s, r.title, { x: r.tx, y: 4.619, w: 1.784, h: 0.421, fontFace: SANS_SB, fontSize: 14, color: r.color, lineSpacingMultiple: LEAD });
        body(s, r.bx, 5.137, 3.194, 0.981,
            'Lorem ipsum dolor sit amet, consectetuer adipiscing porttitor congue massa posuere.', { color: r.bodyColor });
    });

    leafIcon(s, 5.747, 4.753, 0.267, 0.224, WHITE);
    wavesIcon(s, 9.357, 4.753, 0.248, 0.214, PINK);

    brand(s);
    websiteFooter(s);
}

function slide05(pptx) {                                  // Introduce Professional Team
    const s = pptx.addSlide();
    s.background = { color: BG };

    heading(s, 2.086, 1.266, 9.161, 0.841, [
        { text: 'Introduce ' },
        { text: 'Professional', color: PINK },
        { text: ' Team' },
    ], { align: 'center' });

    const TEAM = [
        { name: 'Sebastian Bennett', role: 'Beauty Treatment', x: 2.055, nameX: 1.807, nameW: 2.676, starX: 2.485 },
        { name: 'Rachelle Beaudry', role: 'Beauty Consultant', x: 5.577, nameX: 5.577, nameW: 2.179, starX: 6.007 },
        { name: 'Taylor Alonso', role: 'Administration', x: 9.131, nameX: 9.131, nameW: 2.179, starX: 9.561 },
    ];

    TEAM.forEach(m => {
        photoEllipse(s, m.x, 2.654, 2.179, 2.179, { label: false });
        photoLabel(s, m.x + 1.09, 3.744);
        text(s, m.name, { x: m.nameX, y: 5.205, w: m.nameW, h: 0.37, align: 'center', bold: true, fontFace: SANS_SB, fontSize: 16, color: PINK });
        text(s, m.role, { x: m.x, y: 5.64, w: 2.179, h: 0.303, align: 'center', color: INK_BLACK });
        for (let i = 0; i < 5; i++) {
            s.addShape('star5', { x: m.starX + i * 0.285, y: 6.077, w: 0.18, h: 0.165, fill: { color: PINK } });
        }
    });

    brand(s);
    websiteFooter(s);
}

function slide06(pptx) {                                  // Why You Should Choose Us?
    const s = pptx.addSlide();
    s.background = { color: BG };

    photoBlob(s, 7.492, 0, 4.635, 6.063, SHAPES.roundBottom, [0.5, 0.43]);

    s.addShape('ellipse', { x: 7.543, y: 4.766, w: 1.049, h: 1.049, fill: { color: PINK } });
    leafIcon(s, 7.828, 5.051, 0.48, 0.48, BG);

    heading(s, 1.065, 1.832, 5.08, 1.582, [
        { text: 'Why You Should', breakLine: true },
        { text: 'Choose', color: PINK },
        { text: ' Us?' },
    ]);

    [3.89, 4.852].forEach((y, i) => {
        doubleCheck(s, 1.206, y + 0.182, 0.315, PINK);
        body(s, i === 0 ? 1.799 : 1.806, y, 4.935, 0.678,
            'Lorem ipsum dolor amet, consectetuer adipiscing Maecenas porttitor congue massa dolor.');
    });

    brand(s);
    websiteFooter(s);
}

function slide07(pptx) {                                  // Our Best Service (3 cards)
    const s = pptx.addSlide();
    s.background = { color: BG };

    photoRect(s, 0, 0, 13.333, 4.562, { label: false });
    s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 4.563, fill: { color: WHITE, transparency: 15.3 } });

    brand(s);
    heading(s, 3.399, 1.236, 6.536, 0.841, [
        { text: 'Our Best ' },
        { text: 'Service', color: PINK },
    ], { align: 'center' });
    body(s, 2.285, 2.197, 8.762, 0.678,
        'Lorem ipsum dolor sit amet, consectetuer adipiscing Maecenas porttitor congue massa. ' +
        'Fusce posuere congue massa. Fusce posuere, magna sed pulvinar osuere, posuere', { align: 'center' });

    const SERVICES = [
        { card: 1.772, num: '01.', numX: 2.698, numW: 0.71, title: 'First Service', titleX: 2.15, copyX: 1.992, btn: 2.139 },
        { card: 5.397, num: '02.', numX: 6.323, numW: 0.765, title: 'Second Service', titleX: 5.775, copyX: 5.616, btn: 5.764 },
        { card: 9.021, num: '03.', numX: 9.947, numW: 0.751, title: 'Third Service', titleX: 9.399, copyX: 9.241, btn: 9.388 },
    ];

    SERVICES.forEach(c => {
        s.addShape('roundRect', { x: c.card, y: 3.312, w: 2.54, h: 3.073, rectRadius: 0.215, fill: { color: PINK } });
        text(s, c.num, { x: c.numX, y: 3.474, w: c.numW, h: 0.572, fontFace: SERIF, fontSize: 28, bold: true, color: WHITE });
        text(s, c.title, { x: c.titleX, y: 4.005, w: 1.784, h: 0.421, align: 'center', fontFace: SANS_SB, fontSize: 14, color: WHITE, lineSpacingMultiple: LEAD });
        body(s, c.copyX, 4.477, 2.101, 0.981, 'Lorem ipsum dolor sit amet consectetuer adipiscing.', { align: 'center', color: WHITE });
        s.addShape('roundRect', { x: c.btn, y: 5.631, w: 1.828, h: 0.47, rectRadius: 0.0783, fill: { color: WHITE } });
        text(s, 'Read More', { x: c.btn + 0.134, y: 5.729, w: 1.561, h: 0.303, align: 'center', bold: true, fontFace: SANS_SB, color: INK });
    });

    websiteFooter(s);
}

function slide08(pptx) {                                  // Our Beauty Spa Values
    const s = pptx.addSlide();
    s.background = { color: BG };

    s.addShape('ellipse', { x: -3.788, y: 1.915, w: 10.834, h: 10.834, fill: { color: PINK } });

    // two tilted tablets: silver back shell, black bezel, grey screen
    const BEZEL = [[2.462, 1.031], [4.864, 1.180], [5.979, 4.593], [3.406, 4.608]];
    quad(s, [[1.242, 2.541], [4.300, 2.290], [4.870, 6.240], [2.120, 6.415]], DEVICE_SILVER);
    quad(s, BEZEL, DEVICE_BLACK);
    quad(s, insetQuad(BEZEL, 0.135), PHOTO);
    photoLabel(s, 4.18, 2.85);

    heading(s, 7.452, 1.326, 4.77, 1.582, [
        { text: 'Our Beauty Spa ' },
        { text: 'Values', color: PINK },
    ]);
    body(s, 7.452, 3.045, 5.221, 0.981,
        'Lorem ipsum dolor sit amet, consectetuer adipiscing Maecenas porttitor congue massa. ' +
        'Fusce posuere congue massa. Fusce posuere consectetuer magna.');

    [{ t: 'Value 01', tx: 7.457, bx: 7.453 }, { t: 'Value 02', tx: 10.067, bx: 10.063 }].forEach(v => {
        text(s, v.t, { x: v.tx, y: 4.437, w: 1.784, h: 0.421, fontFace: SANS_SB, fontSize: 14, color: INK, lineSpacingMultiple: LEAD });
        body(s, v.bx, 4.955, 2.61, 0.678, 'Lorem ipsum dolor amet consectetuer porttitor.');
    });

    brand(s);
    websiteFooter(s);
}

function slide09(pptx) {                                  // Providing All You Need
    const s = pptx.addSlide();
    s.background = { color: BG };

    photoEllipse(s, 6.667, 1.23, 5.041, 5.041);

    s.addShape('ellipse', { x: 10.076, y: 3.89, w: 2.534, h: 2.534, fill: { color: PINK } });
    s.addShape('heart', { x: 11.054, y: 4.363, w: 0.572, h: 0.497, fill: { color: WHITE } });
    text(s, 'Lorem ipsum dolor consectetur tempor eiusmod.', {
        x: 10.297, y: 4.987, w: 2.086, h: 0.981, align: 'center', color: WHITE, lineSpacingMultiple: LEAD,
    });

    heading(s, 1.02, 1.716, 6.536, 1.582, [
        { text: 'Providing All', breakLine: true },
        { text: 'You Need', color: PINK },
    ]);
    body(s, 1.02, 3.638, 4.904, 0.981,
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
        'ut labore et dolore magna aliqua. Ut enim ad quis nostrud. ');

    s.addShape('ellipse', { x: 1.086, y: 4.927, w: 0.723, h: 0.723, fill: { color: PINK } });
    s.addShape('star16', { x: 1.288, y: 5.085, w: 0.322, h: 0.322, fill: { color: WHITE } });
    s.addShape('ellipse', { x: 1.325, y: 5.122, w: 0.248, h: 0.248, fill: { color: WHITE } });
    s.addShape('ellipse', { x: 1.360, y: 5.157, w: 0.178, h: 0.178, fill: { color: PINK } });
    s.addShape('roundRect', { x: 1.404, y: 5.205, w: 0.086, h: 0.082, rectRadius: 0.02, fill: { color: WHITE } });
    body(s, 2.069, 4.959, 3.459, 0.678, 'Lorem ipsum dolor sit amet veniam elit, sed eiusmod tempor.');

    brand(s);
    websiteFooter(s);
}

function slide10(pptx) {                                  // Best Quality Beauty Spa
    const s = pptx.addSlide();
    s.background = { color: BG };

    photoRect(s, 0, 0, 5.808, 7.5, { label: false });
    s.addShape('rect', { x: 0, y: 0, w: 5.808, h: 7.5, fill: { color: WHITE, transparency: 15.3 } });

    brand(s);

    photoEllipse(s, 4.576, 1.065, 2.466, 2.466, { label: false });
    photoEllipse(s, 4.576, 3.97, 2.466, 2.466, { label: false });
    photoLabel(s, 5.809, 2.298);

    const FEATURES = [
        { title: 'Best Spa In Town', y: 1.638 },
        { title: 'Complete Facilities', y: 4.546 },
    ];
    FEATURES.forEach(f => {
        text(s, f.title, { x: 1.348, y: f.y, w: 2.797, h: 0.337, bold: true, fontSize: 14, color: INK });
        body(s, 1.348, f.y + 0.409, 2.797, 0.981,
            'Lorem ipsum dolor sit amet, consectetur adipiscing elit, eiusmod tempor');
    });

    heading(s, 7.818, 1.866, 4.848, 1.582, [
        { text: 'Best Quality ', color: INK_BLACK },
        { text: 'Beauty', color: PINK },
        { text: ' Spa', color: INK_BLACK },
    ]);
    body(s, 7.818, 3.795, 4.144, 0.981,
        'Lorem ipsum dolor sit amet, consectetur elit, sed do eiusmod tempor incididunt ut labore ' +
        'et dolore magna enim ad minim.');
    text(s, '2023', { x: 7.816, y: 5.013, w: 1.056, h: 0.572, fontFace: SERIF, fontSize: 28, bold: true, color: PINK });
    body(s, 8.976, 4.975, 3.32, 0.678, 'Lorem ipsum dolor sit consectetur elit, sed do eiusmod tempor.');

    websiteFooter(s);
}

function slide11(pptx) {                                  // Our Best Spa Treatment
    const s = pptx.addSlide();
    s.background = { color: BG };

    s.addShape('rect', { x: 6.667, y: 0, w: 6.667, h: 3.056, fill: { color: PINK } });
    photoRect(s, 1.292, 1.194, 10.75, 3.056, { radius: 0.3147 });

    heading(s, 2.013, 4.724, 4.848, 1.582, [
        { text: 'Our Best Spa', color: INK_BLACK, breakLine: true },
        { text: 'Treatment', color: PINK },
    ]);
    body(s, 7.014, 4.871, 4.935, 0.678,
        'Lorem ipsum dolor amet, consectetuer adipiscing Maecenas porttitor congue massa dolor.');

    [{ label: 'Treatment 01', ix: 7.119, tx: 7.531 }, { label: 'Treatment 02', ix: 9.48, tx: 9.892 }].forEach(t => {
        doubleCheck(s, t.ix, 5.855, 0.315, PINK);
        text(s, t.label, { x: t.tx, y: 5.844, w: 2.15, h: 0.337, bold: true, fontSize: 14, color: INK });
    });

    brand(s);
    websiteFooter(s);
}

function slide12(pptx) {                                  // Take Care Of Your Beautiful
    const s = pptx.addSlide();
    s.background = { color: BG };

    s.addShape('ellipse', { x: 7.684, y: -2.74, w: 7.861, h: 7.861, fill: { color: PINK } });

    // laptop mock-up: black lid, grey screen, silver base with a touchpad lip
    s.addShape('roundRect', { x: 6.750, y: 2.083, w: 5.375, h: 3.584, rectRadius: 0.07, fill: { color: DEVICE_BLACK } });
    photoRect(s, 6.917, 2.292, 5.041, 3.152, { label: false });
    photoLabel(s, 9.438, 3.868);
    s.addShape('roundRect', { x: 6.111, y: 5.667, w: 6.653, h: 0.208, rectRadius: 0.08, fill: { color: DEVICE_SILVER } });
    s.addShape('roundRect', { x: 8.375, y: 5.833, w: 2.125, h: 0.06, rectRadius: 0.03, fill: { color: 'BFC4C3' } });

    heading(s, 1.197, 1.945, 4.848, 1.582, [
        { text: 'Take Care Of', color: INK_BLACK, breakLine: true },
        { text: 'Your ' },
        { text: 'Beautiful', color: PINK },
    ]);
    body(s, 1.197, 3.819, 4.98, 0.981,
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
        'ut labore et dolore magna aliqua. Ut enim ad minim veniam,');

    // clipboard + tick icon
    s.addShape('roundRect', { x: 1.316, y: 5.263, w: 0.349, h: 0.344, rectRadius: 0.05, fill: { color: PINK } });
    s.addShape('rect', { x: 1.4, y: 5.203, w: 0.18, h: 0.08, fill: { color: PINK } });
    checkMark(s, 1.39, 5.33, 0.2, WHITE, 1.5);
    body(s, 1.841, 5.056, 3.942, 0.678,
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod.');

    brand(s);
    websiteFooter(s);
}

function slide13(pptx) {                                  // Our Marketplace Store Here
    const s = pptx.addSlide();
    s.background = { color: BG };

    blob(s, 0, 1.767, 4.717, 5.733, PINK, SHAPES.leftBlob);

    const TABLET = [[0.785, 3.384], [4.592, 1.265], [6.428, 4.090], [2.621, 6.404]];
    quad(s, TABLET, DEVICE_BLACK);
    quad(s, insetQuad(TABLET, 0.130), PHOTO);
    photoLabel(s, 3.607, 3.836);

    heading(s, 7.296, 1.606, 4.865, 1.582, [
        { text: 'Our ', color: INK_BLACK },
        { text: 'Marketplace ', color: PINK },
        { text: 'Store Here', color: INK_BLACK },
    ]);
    body(s, 7.296, 3.517, 4.98, 1.284,
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
        'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ' +
        'ullamco laboris nisi ut aliquip');
    button(s, 7.376, 5.227, 1.828, 0.47, 'More Information');

    brand(s);
    websiteFooter(s);
}

function slide14(pptx) {                                  // Melaty Spa Information
    const s = pptx.addSlide();
    s.background = { color: BG };

    photoBlob(s, 7.138, 0, 6.195, 6.296, SHAPES.topRightPanel, [0.62, 0.62]);

    s.addShape('ellipse', { x: 6.546, y: 1.565, w: 1.221, h: 1.221, fill: { color: PINK } });
    // pencil glyph
    s.addShape('rect', { x: 7.0, y: 2.03, w: 0.28, h: 0.11, rotate: -45, fill: { color: WHITE } });
    s.addShape('triangle', { x: 6.9, y: 2.24, w: 0.13, h: 0.13, rotate: -135, fill: { color: WHITE } });

    heading(s, 1.11, 1.672, 4.024, 1.582, [
        { text: 'Melaty Spa ' },
        { text: 'Information', color: PINK },
    ]);
    body(s, 1.11, 3.492, 5.437, 0.678,
        'Lorem ipsum dolor sit amet, consectetuer adipiscing Maecenas porttitor congue massa posuere.');

    const CONTACTS = [
        { icon: 'globe', value: 'www.yourwebsite.com', ix: 1.203, iy: 4.735, tx: 1.524, ty: 4.625, w: 2.741 },
        { icon: 'mail', value: 'youremail@example', ix: 4.492, iy: 4.784, tx: 4.784, ty: 4.625, w: 2.741 },
        { icon: 'home', value: '123 Street, Country 123', ix: 1.197, iy: 5.423, tx: 1.524, ty: 5.31, w: 2.741 },
        { icon: 'phone', value: '+123 – 456 - 7890', ix: 4.517, iy: 5.459, tx: 4.784, ty: 5.31, w: 2.291 },
    ];
    CONTACTS.forEach(c => {
        contactIcon(s, c.icon, c.ix, c.iy);
        text(s, c.value, { x: c.tx, y: c.ty, w: c.w, h: 0.421, fontSize: 14, color: INK, lineSpacingMultiple: LEAD });
    });

    brand(s);
    websiteFooter(s);
}

/** The four little pink contact glyphs on the information slide. */
function contactIcon(slide, kind, x, y) {
    if (kind === 'globe') {
        slide.addShape('ellipse', { x, y, w: 0.251, h: 0.252, fill: { color: PINK } });
        slide.addShape('ellipse', { x: x + 0.075, y, w: 0.10, h: 0.252, fill: { type: 'none' }, line: { color: WHITE, width: 0.75 } });
        slide.addShape('line', { x, y: y + 0.126, w: 0.251, h: 0, line: { color: WHITE, width: 0.75 } });
    } else if (kind === 'mail') {
        slide.addShape('roundRect', { x, y, w: 0.223, h: 0.154, rectRadius: 0.03, fill: { color: PINK } });
        slide.addShape('line', { x: x + 0.03, y: y + 0.04, w: 0.08, h: 0.06, line: { color: WHITE, width: 1 } });
        slide.addShape('line', { x: x + 0.11, y: y + 0.04, w: 0.08, h: 0.06, line: { color: WHITE, width: 1 }, flipV: true });
    } else if (kind === 'home') {
        slide.addShape('triangle', { x, y, w: 0.264, h: 0.118, fill: { color: PINK } });
        slide.addShape('rect', { x: x + 0.045, y: y + 0.105, w: 0.174, h: 0.108, fill: { color: PINK } });
    } else {
        // handset: crescent band opening towards the upper right
        slide.addShape('blockArc', { x: x - 0.02, y: y - 0.02, w: 0.213, h: 0.213, rotate: 150,
            fill: { color: PINK }, angleRange: [0, 175], arcThicknessRatio: 0.42 });
    }
}

function slide15(pptx) {                                  // Thank You
    const s = pptx.addSlide();
    s.background = { color: BG };

    // full-bleed photo behind a white veil that fades 96% -> 48% opacity, left to right
    const BANDS = 48, BAND_W = 13.333 / BANDS;
    for (let i = 0; i < BANDS; i++) {
        const veil = 0.9569 + (0.4784 - 0.9569) * (i / (BANDS - 1));
        s.addShape('rect', {
            x: i * BAND_W, y: 0, w: BAND_W + 0.02, h: 7.5,
            fill: { color: mix(PHOTO, WHITE, veil) },
        });
    }
    photoLabel(s, 11.3, 6.05);

    brand(s);
    navBar(s);

    text(s, 'Thank You', { x: 0.99, y: 1.922, w: 8.947, h: 1.717, fontFace: DISPLAY, fontSize: 96, bold: true, color: PINK });
    text(s, 'For Your Great Attention.', { x: 1.201, y: 3.505, w: 4.868, h: 0.467, fontFace: SANS_SB, fontSize: 16, color: INK, lineSpacingMultiple: LEAD });
    body(s, 1.201, 4.176, 6.109, 0.981,
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ' +
        'ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco.');
    text(s, 'www.yourwebsite.com', { x: 1.201, y: 5.556, w: 3.81, h: 0.337, fontFace: SANS_SB, fontSize: 14, color: PINK });

    websiteFooter(s);
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
function build() {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'MELATI_16x9', width: 13.333, height: 7.5 });
    pptx.layout = 'MELATI_16x9';
    pptx.author = 'Melati';
    pptx.title = 'Melati - Spa & Beauty Presentation';

    [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
        slide09, slide10, slide11, slide12, slide13, slide14, slide15].forEach(fn => fn(pptx));

    return pptx.writeFile({
        fileName: path.join(__dirname, '15c10ece-54f9-436c-a3b5-81da9aae0a4f_grok_final.pptx'),
    });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
