/*
 * "Mobile Apps Presentation" (BiApps) — 30 slides, 10 x 5.625 in.
 * Rebuilt with pptxgenjs only.  The photographic mock-ups of the original
 * deck (iPhone / iPad / laptop / iMac renders and the small icon PNGs) are
 * re-drawn here with native shapes instead of embedded bitmaps.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ================================================================== *
 * 1. Theme — "Mobile Apps" colour scheme, Poppins / Open Sans
 * ================================================================== */
const ACC1 = 'EC4600';     // accent1
const ACC2 = 'F78400';     // accent2
const ACC3 = 'FA4803';     // accent3 (kicker text)
const WHITE = 'FFFFFF';
const BLACK = '000000';
const DARK = '262626';     // tx1 lum85/off15  — headings
const GRAY25 = '404040';   // tx1 lum75/off25
const GRAY50 = '808080';   // tx1 lum50/off50  — body copy & inactive nav
const GRAY15 = 'D9D9D9';   // bg1 lum85        — empty picture placeholders

const OR_LT = 'FF8B5B';    // accent1 lum60/off40 — end of the house gradient
const OR_LT2 = 'FFB291';   // accent1 lum40/off60 — end of the "button" gradient
const BTN_GRAY = 'E6E6E6';

const FH = 'Poppins';      // major font
const FB = 'Open Sans';    // minor font

const NOLINE = { type: 'none' };
const NOFILL = { type: 'none' };
/* pptxgenjs rewrites the shadow object in place, so hand it a fresh copy each time. */
const shadow = (o = {}) => Object.assign({ type: 'outer', color: BLACK, opacity: 0.15, blur: 40, offset: 0.18, angle: 90 }, o);
const softShadow = () => shadow();
const cardShadow = () => shadow({ color: DARK, opacity: 0.1, blur: 60, offset: 0.25, angle: 45 });

/* ================================================================== *
 * 2. Generic helpers
 * ================================================================== */
function mix(a, b, t) {
    const v = (h, i) => parseInt(h.substr(i, 2), 16);
    const c = i => Math.round(v(a, i) + (v(b, i) - v(a, i)) * t).toString(16).padStart(2, '0');
    return (c(0) + c(2) + c(4)).toUpperCase();
}

/**
 * The deck's house gradient (accent1 -> lighter accent1, 45 degrees).
 * pptxgenjs cannot emit gradient fills, so the real shape is drawn once in the
 * mid tone and a grid of small rectangles is laid inside it to shade it.
 *   r      corner radius in inches: a number, or {tl,tr,bl,br}
 *   shape  optional preset for the backing shape (defaults to rect / roundRect)
 *   dir    gradient direction: 45 (default, dark top-left), 135 or 225 for the
 *          shapes the original rotates by 90 / 180 degrees
 */
const CELL = 0.22;     // shading cell size in inches — fine enough to read as smooth
const OVERLAP = 0.006; // cells overlap a hair so no hairline seams show through

function gradFill(s, o) {
    const c1 = o.c1 || ACC1, c2 = o.c2 || OR_LT, dir = o.dir || 45;
    const ramp = (fx, fy) => dir === 135 ? (1 - fx + fy) / 2 : dir === 225 ? (2 - fx - fy) / 2 : (fx + fy) / 2;
    const R = Object.assign({ tl: 0, tr: 0, bl: 0, br: 0 }, typeof o.r === 'number' ? { tl: o.r, tr: o.r, bl: o.r, br: o.r } : o.r);
    const caps = [R.tl, R.tr, R.bl, R.br], uniform = caps.every(v => v === caps[0]);

    // 1. one backing shape carries the true silhouette
    const base = { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: mix(c1, c2, 0.5) }, line: NOLINE };
    const sideways = { x: o.x + (o.w - o.h) / 2, y: o.y + (o.h - o.w) / 2, w: o.h, h: o.w };
    if (o.shape) s.addShape(o.shape, base);
    else if (uniform) s.addShape(caps[0] ? 'roundRect' : 'rect', Object.assign(base, caps[0] ? { rectRadius: caps[0] } : {}));
    else if (R.tl && R.tr) s.addShape('round2SameRect', base);
    else if (R.tr && R.br) s.addShape('round2SameRect', Object.assign(base, sideways, { rotate: 90 }));
    else if (R.tl && R.bl) s.addShape('round2SameRect', Object.assign(base, sideways, { rotate: 270 }));
    else s.addShape('round1Rect', Object.assign(base, { rectRadius: R.tr || R.tl || R.br || R.bl }));

    // 2. how far a rounded corner eats into the top / bottom edge at offset dx
    const inset = (dx, rL, rR) => {
        const rr = dx < rL ? rL : dx > o.w - rR ? rR : 0;
        if (!rr) return 0;
        const d = dx < rL ? rL - dx : rR - (o.w - dx);
        return rr - Math.sqrt(Math.max(0, rr * rr - d * d));
    };
    const top = dx => o.y + inset(dx, R.tl, R.tr);
    const bot = dx => o.y + o.h - inset(dx, R.bl, R.br);

    // 3. column edges — dense through the corner arcs, CELL-sized across the middle
    const capL = Math.max(R.tl, R.bl), capR = Math.max(R.tr, R.br);
    const edges = [0];
    const span = (from, to, k) => { for (let i = 1; i <= k; i++) edges.push(from + ((to - from) * i) / k); };
    if (capL) span(0, capL, Math.max(10, Math.round(capL * 40)));
    span(capL, o.w - capR, Math.max(1, Math.round((o.w - capL - capR) / CELL)));
    if (capR) span(o.w - capR, o.w, Math.max(10, Math.round(capR * 40)));

    // 4. shade every column, keeping each cell strictly inside the silhouette
    for (let i = 0; i < edges.length - 1; i++) {
        const x0 = edges[i], x1 = edges[i + 1];
        if (x1 - x0 < 0.002) continue;
        const t = (x0 + x1) / (2 * o.w);
        const y0 = Math.max(top(x0), top(x1)), y1 = Math.min(bot(x0), bot(x1));
        if (y1 - y0 < 0.012) continue;
        const n = Math.max(1, Math.round((y1 - y0) / CELL)), dy = (y1 - y0) / n;
        for (let j = 0; j < n; j++) {
            const y = y0 + dy * j;
            s.addShape('rect', {
                x: o.x + x0, y, w: x1 - x0 + (i < edges.length - 2 ? OVERLAP : 0),
                h: dy + (j < n - 1 ? OVERLAP : 0),
                fill: { color: mix(c1, c2, ramp(t, (y + dy / 2 - o.y) / o.h)) }, line: NOLINE,
            });
        }
    }
}

/** Gradient-filled circle. */
const gradCircle = (s, o) => gradFill(s, Object.assign({}, o, { shape: 'ellipse', r: o.w / 2 }));

/* ---- the shared page chrome: nav links + BiApps wordmark ---- */
const NAV_LABELS = ['Home', 'About Us', 'Service', 'Team'];
const NAV_X = [0.786, 1.884, 2.982, 4.08];

function addNav(s, o = {}) {
    const xs = o.xs || NAV_X;
    NAV_LABELS.forEach((label, i) => s.addText(label, {
        x: xs[i], y: 0.479, w: 0.877, h: 0.252, valign: 'top',
        fontFace: FH, fontSize: 9, bold: i === 0,
        color: o.color || (i === 0 ? DARK : GRAY50),
    }));
}

function addLogo(s, o = {}) {
    s.addText(o.color ? [{ text: 'BiApps' }]
        : [{ text: 'Bi', options: { color: DARK } }, { text: 'Apps', options: { color: ACC1 } }], {
        x: o.x === undefined ? 8.167 : o.x, y: 0.454, w: 1.047, h: 0.278, valign: 'top',
        align: o.align || 'right', fontFace: FH, fontSize: 10.5, bold: true, color: o.color,
    });
}

function addChrome(s, o) { addNav(s, o); addLogo(s, o); }

/* ---- text styles lifted from the deck's list styles ---- */
const titleStyle = (o = {}) => Object.assign({ fontFace: FH, fontSize: 27, bold: true, color: DARK, valign: 'top' }, o);
const kickerStyle = (o = {}) => Object.assign({
    fontFace: FH, fontSize: 8.25, bold: true, charSpacing: 0.98, color: ACC3,
    align: 'justify', lineSpacingMultiple: 1.5, valign: 'top',
}, o);
const labelStyle = (o = {}) => Object.assign({
    fontFace: FH, fontSize: 8.25, bold: true, charSpacing: 1.3, color: GRAY25, valign: 'top',
}, o);
const bodyStyle = (o = {}) => Object.assign({
    fontFace: FB, fontSize: 7.88, color: GRAY50, lineSpacingMultiple: 1.3, valign: 'top',
}, o);
const cardHead = (o = {}) => Object.assign({ fontFace: FB, fontSize: 9, bold: true, color: BLACK, valign: 'top' }, o);
const numStyle = (o = {}) => Object.assign({ fontFace: FH, fontSize: 21, bold: true, color: ACC1, valign: 'top' }, o);

/* ================================================================== *
 * 3. Device mock-ups (stand-ins for the deck's photographs)
 * ================================================================== */
/** Hollow phone bezel: dark frame + metal edge highlight, screen stays transparent. */
function phone(s, x, y, w, h, o = {}) {
    const t = w * 0.062, rot = o.rotate;
    s.addShape('roundRect', {
        x: x + t / 2, y: y + t / 2, w: w - t, h: h - t, rotate: rot,
        rectRadius: w * 0.108, fill: NOFILL, line: { color: '2C2C2C', width: t * 72 },
    });
    s.addShape('roundRect', {                                     // brushed-metal edge
        x: x + t * 0.24, y: y + t * 0.24, w: w - t * 0.48, h: h - t * 0.48, rotate: rot,
        rectRadius: w * 0.112, fill: NOFILL, line: { color: 'C4C4C4', width: t * 15 },
    });
    s.addShape('roundRect', { x, y, w, h, rotate: rot, rectRadius: w * 0.115, fill: NOFILL, line: { color: '8C8C8C', width: 0.75 } });
    if (!rot) {                                                   // notch
        s.addShape('roundRect', { x: x + w * 0.31, y: y + t * 0.8, w: w * 0.38, h: h * 0.024, rectRadius: h * 0.012, fill: { color: '2C2C2C' }, line: NOLINE });
    }
}

/** iPad: rounded black frame, transparent screen. */
function tablet(s, x, y, w, h) {
    s.addShape('roundRect', { x: x + 0.045, y: y + 0.045, w: w - 0.09, h: h - 0.09, rectRadius: 0.16, fill: NOFILL, line: { color: '121212', width: 6.5 } });
    s.addShape('roundRect', { x, y, w, h, rectRadius: 0.2, fill: NOFILL, line: { color: '8A8A8A', width: 1 } });
}

/** Open laptop seen head-on: black screen frame over a silver wedge base. */
function laptop(s, x, y, w, h) {
    s.addShape('rect', { x, y, w, h, fill: NOFILL, line: { color: '111111', width: 8 } });
    s.addShape('trapezoid', { x: x - 0.34, y: y + h + 0.03, w: w + 0.68, h: 0.17, rotate: 180, fill: { color: 'D3D5D8' }, line: NOLINE });
    s.addShape('rect', { x: x - 0.06, y: y + h - 0.02, w: w + 0.12, h: 0.06, fill: { color: 'EDEEEF' }, line: NOLINE });
}

/** iMac at a slight tilt: panel on a tapered neck and a flat foot. */
function imac(s, x, y, w, h) {
    s.addShape('trapezoid', { x: x + w * 0.33, y: y + h - 0.04, w: w * 0.24, h: 0.62, fill: { color: 'F3F3F3' }, line: NOLINE });
    s.addShape('roundRect', { x: x + w * 0.18, y: y + h + 0.54, w: w * 0.5, h: 0.1, rectRadius: 0.05, fill: { color: 'ECECEC' }, line: NOLINE });
    s.addShape('roundRect', { x, y, w, h, rotate: -5, rectRadius: 0.08, fill: { color: 'F4F4F2' }, line: { color: 'E6E6E4', width: 1 } });
    s.addShape('rect', { x: x + 0.13, y: y + 0.13, w: w - 0.26, h: h - 0.34, rotate: -5, fill: { color: 'E9EAE8' }, line: NOLINE });
}

/** Diamond BiApps logo + wordmark + "Try Now" pill, i.e. the in-app splash. */
function appScreen(s, o) {
    const d = o.logo;
    s.addShape('diamond', { x: d.x, y: d.y, w: d.w, h: d.h, fill: NOFILL, line: { color: WHITE, width: d.w * 7 } });
    const bw = d.w * 0.1;
    [-1, 0, 1].forEach(k => s.addShape('rect', {
        x: d.x + d.w / 2 - bw / 2 + k * d.w * 0.17, y: d.y + d.h * 0.4,
        w: bw, h: d.h * (k === 0 ? 0.32 : 0.24), fill: { color: WHITE }, line: NOLINE,
    }));
    s.addText('BiApps', Object.assign({ fontFace: FH, color: WHITE, align: 'center', valign: 'top' }, o.word));
    tryPill(s, o.btn.x, o.btn.y, o.btn.w, o.btn.h, o.btn.fontSize, WHITE);
}

function tryPill(s, x, y, w, h, fontSize, fill) {
    s.addText('Try Now', {
        x, y, w, h, shape: 'roundRect', rectRadius: h / 2, fill: { color: fill }, line: NOLINE,
        fontFace: FB, fontSize, bold: true, color: BLACK, align: 'center', valign: 'middle',
    });
}

/** Un-filled "Click icon to add picture" placeholders left in the original. */
function picPlaceholder(s, x, y, w, h) {
    s.addShape('roundRect', { x, y, w, h, rectRadius: 0.09, fill: { color: GRAY15 }, line: NOLINE });
    s.addShape('rect', { x: x + w / 2 - 0.32, y: y + h / 2 - 0.24, w: 0.64, h: 0.48, fill: { color: WHITE }, line: { color: '9E9E9E', width: 0.75 } });
    s.addShape('triangle', { x: x + w / 2 - 0.22, y: y + h / 2 - 0.04, w: 0.44, h: 0.26, fill: { color: '86B4DE' }, line: NOLINE });
    s.addShape('ellipse', { x: x + w / 2 - 0.18, y: y + h / 2 - 0.15, w: 0.1, h: 0.1, fill: { color: 'F2C14B' }, line: NOLINE });
    s.addText('Click icon to add picture', { x: x - 1.1, y: y - 0.06, w: w + 2.2, h: 0.3, align: 'center', valign: 'top', fontFace: FB, fontSize: 13.5, color: '333333' });
}

/* ---- pictograms used inside the round badges (all native shapes) ---- */
function icon(s, kind, x, y, sz, color, bg) {
    const F = { color }, L = NOLINE;
    if (kind === 'mail') {
        s.addShape('rect', { x, y: y + sz * 0.18, w: sz, h: sz * 0.62, fill: F, line: L });
        s.addShape('line', { x, y: y + sz * 0.18, w: sz / 2, h: sz * 0.33, line: { color: bg, width: 1.25 } });
        s.addShape('line', { x: x + sz / 2, y: y + sz * 0.51, w: sz / 2, h: sz * 0.33, flipV: true, line: { color: bg, width: 1.25 } });
    } else if (kind === 'users') {
        s.addShape('ellipse', { x: x + sz * 0.04, y, w: sz * 0.34, h: sz * 0.34, fill: F, line: L });
        s.addShape('ellipse', { x: x + sz * 0.58, y, w: sz * 0.34, h: sz * 0.34, fill: F, line: L });
        s.addShape('round2SameRect', { x, y: y + sz * 0.44, w: sz * 0.42, h: sz * 0.46, fill: F, line: L });
        s.addShape('round2SameRect', { x: x + sz * 0.54, y: y + sz * 0.44, w: sz * 0.42, h: sz * 0.46, fill: F, line: L });
    } else if (kind === 'person') {
        s.addShape('ellipse', { x: x + sz * 0.28, y, w: sz * 0.44, h: sz * 0.44, fill: F, line: L });
        s.addShape('round2SameRect', { x: x + sz * 0.1, y: y + sz * 0.52, w: sz * 0.8, h: sz * 0.44, fill: F, line: L });
    } else if (kind === 'cart') {
        s.addShape('trapezoid', { x, y: y + sz * 0.16, w: sz * 0.92, h: sz * 0.42, rotate: 180, fill: F, line: L });
        s.addShape('ellipse', { x: x + sz * 0.14, y: y + sz * 0.64, w: sz * 0.2, h: sz * 0.2, fill: F, line: L });
        s.addShape('ellipse', { x: x + sz * 0.58, y: y + sz * 0.64, w: sz * 0.2, h: sz * 0.2, fill: F, line: L });
    } else if (kind === 'check') {              // two rotated bars form the tick
        s.addShape('rect', { x: x + sz * 0.06, y: y + sz * 0.52, w: sz * 0.34, h: sz * 0.15, rotate: 45, fill: F, line: L });
        s.addShape('rect', { x: x + sz * 0.28, y: y + sz * 0.4, w: sz * 0.72, h: sz * 0.15, rotate: -52, fill: F, line: L });
    } else if (kind === 'chevron') {
        s.addShape('rect', { x: x + sz * 0.22, y: y + sz * 0.18, w: sz * 0.52, h: sz * 0.12, rotate: 48, fill: F, line: L });
        s.addShape('rect', { x: x + sz * 0.22, y: y + sz * 0.68, w: sz * 0.52, h: sz * 0.12, rotate: -48, fill: F, line: L });
    } else if (kind === 'phone') {              // handset: rotated bar + two ear pieces
        s.addShape('roundRect', { x: x + sz * 0.1, y: y + sz * 0.34, w: sz * 0.8, h: sz * 0.3, rotate: -40, rectRadius: sz * 0.14, fill: F, line: L });
        s.addShape('ellipse', { x: x + sz * 0.02, y: y + sz * 0.02, w: sz * 0.34, h: sz * 0.34, fill: F, line: L });
        s.addShape('ellipse', { x: x + sz * 0.62, y: y + sz * 0.62, w: sz * 0.34, h: sz * 0.34, fill: F, line: L });
    } else if (kind === 'web') {                // laptop with a globe on the screen
        s.addShape('rect', { x, y, w: sz, h: sz * 0.72, fill: NOFILL, line: { color, width: 1.25 } });
        s.addShape('ellipse', { x: x + sz * 0.3, y: y + sz * 0.16, w: sz * 0.4, h: sz * 0.4, fill: F, line: L });
        s.addShape('rect', { x: x - sz * 0.12, y: y + sz * 0.78, w: sz * 1.24, h: sz * 0.1, fill: F, line: L });
    } else if (kind === 'pin') {
        s.addShape('teardrop', { x, y, w: sz * 0.78, h: sz * 0.78, rotate: 135, fill: F, line: L });
        s.addShape('ellipse', { x: x + sz * 0.25, y: y + sz * 0.19, w: sz * 0.28, h: sz * 0.28, fill: { color: bg || WHITE }, line: L });
    }
}

/* ================================================================== *
 * 4. Slides
 * ================================================================== */
const LOREM_A = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ';
const LOREM_B = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus vel euismod leo. ';
const DEV_TEXT = LOREM_A + 'Fusce posuere, magna sed.';

function slide01(s) {
    // full-width panel, only the top-right corner rounded (round1Rect, r=0.715)
    gradFill(s, { x: 0, y: 1.486, w: 9.48, h: 4.139, r: { tr: 0.715 } });
    phone(s, 6.567, 1.076, 2.07, 4.112);
    s.addText([{ text: 'Mobile ', options: { bold: true } }, { text: 'Apps Presentation', options: { bold: false } }],
        { x: 1.049, y: 2.017, w: 4.696, h: 1.616, fontFace: FH, fontSize: 45, color: WHITE, valign: 'top' });
    tryPill(s, 1.156, 4.579, 1.27, 0.354, 10.5, BTN_GRAY);
    s.addText(LOREM_A + 'Fusce posuere, magna sed', bodyStyle({ x: 1.034, y: 3.756, w: 4.202, h: 0.431, color: WHITE }));
    addNav(s, { xs: [5.634, 6.732, 7.83, 8.928] });
    addLogo(s, { x: 0.75, align: 'left' });
    appScreen(s, {
        logo: { x: 7.2695, y: 1.889, w: 0.664, h: 0.701 },
        word: { x: 6.695, y: 2.8775, w: 1.8135, h: 0.4545, fontSize: 21 },
        btn: { x: 7.0475, y: 4.133, w: 1.1085, h: 0.309, fontSize: 10.5 },
    });
}

function slide02(s) {
    gradCircle(s, { x: 1.151, y: 1.723, w: 3.123, h: 3.123 });
    phone(s, 1.677, 1.244, 2.07, 4.112);
    s.addText('Discover The World Of App Mobile', titleStyle({ x: 5.294, y: 1.767, w: 4.066, h: 1.01 }));
    s.addText([
        { text: LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.', options: { breakLine: true } },
        { text: ' ', options: { breakLine: true } },
        { text: 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.' },
    ], bodyStyle({ x: 5.294, y: 3.496, w: 3.724, h: 1.121 }));
    s.addText('DESCRIPTION HERE', labelStyle({ x: 5.294, y: 3.096, w: 1.635, h: 0.29 }));
    appScreen(s, {
        logo: { x: 2.3795, y: 2.065, w: 0.664, h: 0.701 },
        word: { x: 1.805, y: 3.0535, w: 1.8135, h: 0.4545, fontSize: 21 },
        btn: { x: 2.1575, y: 4.309, w: 1.1085, h: 0.309, fontSize: 10.5 },
    });
    addChrome(s);
}

function slide03(s) {
    phone(s, 6.105, 1.049, 2.864, 5.689);
    s.addShape('roundRect', {
        x: 0.988, y: 3.87, w: 3.635, h: 0.786, rectRadius: 0.13, fill: { color: WHITE }, line: NOLINE,
        shadow: shadow({ opacity: 0.06, blur: 50, offset: 0.11, angle: 70 }),
    });
    s.addShape('line', { x: 1.902, y: 4.014, w: 0, h: 0.497, line: { color: GRAY15, width: 1 } });
    s.addText(LOREM_B, bodyStyle({ x: 2.012, y: 4.074, w: 2.437, h: 0.431 }));
    s.addShape('ellipse', { x: 1.126, y: 3.979, w: 0.566, h: 0.568, fill: { color: ACC1 }, line: NOLINE, shadow: softShadow() });
    icon(s, 'mail', 1.264, 4.19, 0.29, WHITE, ACC1);
    s.addText('Mobile Application Development ', titleStyle({ x: 0.964, y: 1.35, w: 4.036, h: 1.01 }));
    s.addText(LOREM_B + 'Donec commodo et urna ac semper.', bodyStyle({ x: 0.964, y: 3.079, w: 3.724, h: 0.431 }));
    s.addText('DESCRIPTION HERE', labelStyle({ x: 0.964, y: 2.679, w: 1.635, h: 0.29 }));
    addChrome(s);
    appScreen(s, {
        logo: { x: 7.0115, y: 1.917, w: 1.0665, h: 1.126 },
        word: { x: 6.506, y: 3.3728, w: 2.0775, h: 0.555, fontSize: 27 },
        btn: { x: 6.9028, y: 4.8105, w: 1.2698, h: 0.354, fontSize: 10.5 },
    });
}

function slide04(s) {
    gradFill(s, { x: 0.72, y: 2.812, w: 8.56, h: 2.812, r: { tl: 0.469, tr: 0.469 } });
    phone(s, 1.712, 1.03, 2.864, 5.689);
    s.addText([{ text: 'Mobile Apps', options: { breakLine: true } }, { text: 'Service' }], titleStyle({ x: 5.341, y: 1.479, w: 3.305, h: 1.01 }));
    s.addText('MOCKUP SECTION', kickerStyle({ x: 5.357, y: 1.082, w: 1.8, h: 0.29 }));
    s.addText('Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar osuere, magna, magna sed pulvinar osuere, posuere, magna',
        { x: 5.341, y: 3.768, w: 3.305, h: 0.912, fontFace: FB, fontSize: 8.25, color: WHITE, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText('OUR DESCRIPTION HERE', labelStyle({ x: 5.357, y: 3.34, w: 2.267, h: 0.29, color: WHITE }));
    addChrome(s);
    appScreen(s, {
        logo: { x: 2.6105, y: 1.917, w: 1.0665, h: 1.126 },
        word: { x: 2.105, y: 3.3728, w: 2.0775, h: 0.555, fontSize: 27 },
        btn: { x: 2.5018, y: 4.8105, w: 1.2698, h: 0.354, fontSize: 10.5 },
    });
}

function slide05(s) {
    // a round2SameRect rotated 90 deg: covers the whole slide bar a rounded right edge
    gradFill(s, { x: 0, y: 0.661, w: 9.474, h: 4.304, r: { tr: 0.35, br: 0.35 }, dir: 135 });
    phone(s, 5.341, 1.271, 2.864, 5.689);
    s.addText('App Uses AI to Enhance Compatibility', titleStyle({ x: 1.017, y: 1.718, w: 4.191, h: 1.464, color: WHITE }));
    s.addText('MOCKUP SECTION', kickerStyle({ x: 1.033, y: 1.321, w: 1.8, h: 0.29, color: WHITE }));
    [1.017, 3.066].forEach(x => s.addText('Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor',
        { x, y: 3.727, w: 1.803, h: 0.704, fontFace: FB, fontSize: 8.25, color: WHITE, align: 'justify', lineSpacingMultiple: 1.5, valign: 'top' }));
    s.addText('OUR DESCRIPTION HERE', labelStyle({ x: 1.033, y: 3.299, w: 1.972, h: 0.29, color: WHITE }));
    s.addText('BiApps', { x: 5.734, y: 2.282, w: 2.077, h: 0.555, fontFace: FH, fontSize: 27, color: WHITE, align: 'center', valign: 'top' });

    // the little "Performance" analytics card that floats over the phone
    s.addShape('roundRect', { x: 5.737, y: 3.494, w: 2.077, h: 1.728, rectRadius: 0.26, fill: { color: WHITE }, line: NOLINE, shadow: softShadow() });
    s.addText('Performance', { x: 6.043, y: 3.578, w: 1.466, h: 0.253, fontFace: FH, fontSize: 9, color: '031134', align: 'center', valign: 'top' });
    s.addChart('line', [{ name: 'Uptrend', labels: ['Feb', 'Mar', 'Apr', 'Jun', 'Jul', 'Agu'], values: [42, 26, 22, 46, 30, 62] }], {
        x: 5.83, y: 3.82, w: 1.95, h: 1.33, chartColors: [ACC1], lineSize: 2, lineDataSymbol: 'none',
        showLegend: false, showTitle: false, valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20,
        catAxisLabelFontSize: 4.5, valAxisLabelFontSize: 4.5, catAxisLabelFontFace: FH, valAxisLabelFontFace: FH,
        catAxisLabelColor: GRAY50, valAxisLabelColor: GRAY50, catAxisLineShow: false, valAxisLineShow: false,
        valGridLine: { style: 'solid', size: 0.25, color: 'E8E8E8' }, catGridLine: { style: 'solid', size: 0.25, color: 'E8E8E8' },
        plotArea: { fill: { color: WHITE } },
    });
    s.addText('Uptrend', {
        x: 6.897, y: 4.059, w: 0.501, h: 0.194, shape: 'roundRect', rectRadius: 0.03, fill: { color: '031134' },
        line: NOLINE, fontFace: FH, fontSize: 5.25, color: WHITE, align: 'center', valign: 'middle',
    });
}

function slide06(s) {
    gradCircle(s, { x: 5.23, y: -0.508, w: 6.641, h: 6.641 });
    picPlaceholder(s, 4.685, 1.255, 1.574, 3.343);
    picPlaceholder(s, 6.938, 1.006, 1.807, 3.837);
    phone(s, 6.804, 0.874, 2.078, 4.127);
    phone(s, 4.567, 1.14, 1.81, 3.596);
    s.addText([{ text: 'Mobile Apps', options: { breakLine: true } }, { text: 'Service' }], titleStyle({ x: 1.039, y: 1.479, w: 3.305, h: 1.01 }));
    s.addText('MOCKUP SECTION', kickerStyle({ x: 1.054, y: 1.082, w: 1.8, h: 0.29 }));
    [['Service For Company', 2.813, 3.101, LOREM_B],
     ['Service For Personal', 3.794, 4.083, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit ipsum dolor sit amet, consectetur ']]
        .forEach(([head, yh, yb, body]) => {
            s.addText(head, labelStyle({ x: 1.363, y: yh, w: 1.749, h: 0.29, charSpacing: 0 }));
            s.addText(body, bodyStyle({ x: 1.363, y: yb, w: 2.85, h: 0.431 }));
            icon(s, 'check', 1.06, yh + 0.036, 0.25, ACC1);
        });
    addNav(s);
    addLogo(s, { color: WHITE });
    appScreen(s, {
        logo: { x: 7.5325, y: 1.661, w: 0.664, h: 0.701 },
        word: { x: 6.958, y: 2.6495, w: 1.8135, h: 0.4545, fontSize: 21 },
        btn: { x: 7.3105, y: 3.905, w: 1.1085, h: 0.309, fontSize: 10.5 },
    });
}

function slide07(s) {
    gradFill(s, { x: 0.525, y: 2.812, w: 5.01, h: 2.812, r: { tl: 0.469, tr: 0.469 }, c2: mix(ACC1, OR_LT, 0.6) });
    s.addText('Personal Security App', titleStyle({ x: 6.049, y: 1.644, w: 3.305, h: 1.01 }));
    s.addText('MOCKUP SECTION', kickerStyle({ x: 6.064, y: 1.247, w: 1.8, h: 0.29 }));
    s.addText(LOREM_B + 'Donec commodo et urn', bodyStyle({ x: 6.036, y: 3.394, w: 3.174, h: 0.431 }));
    s.addText('DESCRIPTION HERE', labelStyle({ x: 6.036, y: 2.994, w: 1.635, h: 0.29 }));
    [4.043, 4.457].forEach(y => {
        s.addText('Service For Company', labelStyle({ x: 6.343, y, w: 1.749, h: 0.29, charSpacing: 0 }));
        icon(s, 'check', 6.04, y + 0.036, 0.25, ACC1);
    });
    addChrome(s);
}

function slide08(s) {
    gradFill(s, { x: 5, y: 0, w: 5, h: 5.625 });
    s.addText('Boosting Productivity', titleStyle({ x: 1.039, y: 1.479, w: 3.305, h: 1.01 }));
    s.addText('MOCKUP SECTION', kickerStyle({ x: 1.054, y: 1.082, w: 1.8, h: 0.29 }));
    [[5.574, 2.34], [7.465, 2.34], [5.574, 4.743], [7.465, 4.743]].forEach(([x, y]) =>
        s.addText('OUR DESCRIPTION', kickerStyle({ x, y, w: 1.8, h: 0.29, color: WHITE, align: 'center' })));
    s.addText(['Far behind the word ', 'Far countries ', 'Far behind the word ', 'Far countries']
        .map(t => ({ text: t, options: { breakLine: true, bullet: { characterCode: '2713', indent: 10 } } })),
        { x: 1.057, y: 3.686, w: 2.262, h: 0.988, fontFace: FH, fontSize: 9, bold: true, color: GRAY25, lineSpacingMultiple: 1.5, valign: 'top' });
    s.addText(LOREM_B + 'Donec commodo et urn', bodyStyle({ x: 1.043, y: 3.039, w: 3.174, h: 0.431 }));
    s.addText('DESCRIPTION HERE', labelStyle({ x: 1.043, y: 2.638, w: 1.635, h: 0.29 }));
}

function slide09(s) {
    gradFill(s, { x: 0, y: 0, w: 3.602, h: 3.648 });
    s.addText('Educational App Makes Learning', titleStyle({ x: 5.902, y: 1.479, w: 3.598, h: 1.01 }));
    s.addText('MOCKUP SECTION', kickerStyle({ x: 5.918, y: 1.082, w: 1.8, h: 0.29 }));
    [['Development 01', ACC1, 2.651, 2.96, 'mail', WHITE, ACC1],
     ['Development 02', ACC2, 3.724, 4.032, 'users', ACC2, WHITE]]
        .forEach(([head, col, yh, yb, ic, icCol, badge]) => {
            s.addText(head, { x: 6.66, y: yh, w: 1.967, h: 0.281, fontFace: FB, fontSize: 9, bold: true, color: col, lineSpacingMultiple: 1.3, valign: 'top' });
            s.addText(DEV_TEXT, bodyStyle({ x: 6.66, y: yb, w: 2.783, h: 0.604 }));
            s.addShape('ellipse', { x: 5.958, y: yh + 0.047, w: 0.566, h: 0.568, fill: { color: badge }, line: NOLINE, shadow: softShadow() });
            icon(s, ic, 6.098, yh + 0.215, 0.28, icCol, badge);
        });
}

function slide10(s) {
    gradFill(s, { x: 0.72, y: 3.398, w: 8.56, h: 2.227, r: { tl: 0.371, tr: 0.371 } });
    s.addText('App Helps You Split Expenses', titleStyle({ x: 1.039, y: 1.688, w: 3.305, h: 1.01 }));
    s.addText('MOCKUP SECTION', kickerStyle({ x: 1.054, y: 1.29, w: 1.8, h: 0.29 }));
    [['Development 01', 2.002, 1.3], ['Development 02', 6.063, 5.362]].forEach(([head, xt, xi]) => {
        s.addText(head, { x: xt, y: 4.079, w: 1.967, h: 0.281, fontFace: FB, fontSize: 9, bold: true, color: WHITE, lineSpacingMultiple: 1.3, valign: 'top' });
        s.addText(DEV_TEXT, bodyStyle({ x: xt, y: 4.388, w: 2.783, h: 0.604, color: WHITE }));
        s.addShape('ellipse', { x: xi, y: 4.127, w: 0.566, h: 0.568, fill: { color: WHITE }, line: NOLINE, shadow: softShadow() });
        icon(s, 'users', xi + 0.14, 4.295, 0.29, ACC2);
    });
    addChrome(s);
}

function slide11(s) {
    gradFill(s, { x: 0.72, y: 2.363, w: 8.56, h: 2.764, r: 0.284 });
    phone(s, 3.688, 1.727, 2.624, 5.214);
    s.addText('Mobile Apps Service', titleStyle({ x: 2.638, y: 1.021, w: 4.724, h: 0.555, align: 'center' }));
    [['Development 03', 6.753, 2.783, 5.977, 'left'], ['Development 04', 6.753, 3.935, 5.977, 'left'],
     ['Development 01', 1.142, 2.783, 3.456, 'right'], ['Development 02', 1.142, 3.935, 3.456, 'right']]
        .forEach(([head, xt, y, xi, align]) => {
            s.addText(head, { x: xt, y, w: 1.967, h: 0.281, align, fontFace: FB, fontSize: 9, bold: true, color: WHITE, lineSpacingMultiple: 1.3, valign: 'top' });
            s.addText(LOREM_A, bodyStyle({ x: xt, y: y + 0.308, w: 2.112, h: 0.604, color: WHITE, align }));
            s.addShape('ellipse', { x: xi, y: y + 0.047, w: 0.566, h: 0.568, fill: { color: WHITE }, line: NOLINE, shadow: softShadow() });
            icon(s, 'users', xi + 0.14, y + 0.215, 0.29, ACC2);
        });
    addChrome(s);
}

function slide12(s) {
    // wedge: 5.479 in wide at the top, tapering to 3.153 in at the bottom
    gradFill(s, { x: 0, y: 0, w: 5.479, h: 5.625, c2: mix(ACC1, OR_LT, 0.62) });
    s.addShape('rtTriangle', { x: 3.153, y: 0, w: 2.326, h: 5.625, flipH: true, fill: { color: WHITE }, line: NOLINE });
    picPlaceholder(s, -0.92, 1.32, 1.397, 2.966);
    picPlaceholder(s, 1.026, 1.046, 1.653, 3.51);
    picPlaceholder(s, 3.214, 1.32, 1.397, 2.966);
    phone(s, 3.11, 1.218, 1.606, 3.19);
    appScreen(s, {
        logo: { x: 3.6108, y: 1.641, w: 0.564, h: 0.5955 },
        word: { x: 2.986, y: 2.5253, w: 1.8135, h: 0.3533, fontSize: 15 },
        btn: { x: 3.3385, y: 3.78, w: 1.1085, h: 0.309, fontSize: 8.25 },
    });
    phone(s, 0.904, 0.925, 1.901, 3.775);
    phone(s, -1.024, 1.218, 1.606, 3.19);
    s.addShape('ellipse', { x: 2.139, y: 4.073, w: 0.917, h: 0.917, fill: { color: WHITE }, line: NOLINE, shadow: shadow({ color: '282828', opacity: 0.21, blur: 18, offset: 0.06, angle: 90 }) });
    s.addShape('ellipse', { x: 2.335, y: 4.27, w: 0.524, h: 0.524, fill: NOFILL, line: { color: OR_LT, width: 1.5, dashType: 'sysDash' } });
    icon(s, 'check', 2.43, 4.35, 0.34, ACC1);
    s.addText(LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero', bodyStyle({ x: 5.877, y: 2.624, w: 3.389, h: 0.604 }));
    s.addText('Mobile App User Download', titleStyle({ x: 5.877, y: 1.468, w: 3.389, h: 1.01 }));
    [['mail', 5.943, ACC1], ['users', 6.815, ACC2], ['cart', 7.687, ACC1]].forEach(([kind, x, fill]) => {
        s.addShape('ellipse', { x, y: 3.453, w: 0.566, h: 0.568, fill: { color: fill }, line: NOLINE, shadow: softShadow() });
        icon(s, kind, x + 0.14, 3.62, 0.28, WHITE, fill);
    });
    addNav(s, { color: WHITE });
    addLogo(s);
}

function slide13(s) {
    phone(s, 0.39, 1.32, 3.654, 7.258, { rotate: 19.28 });
    // the reference group scales a 3.016 x 0.275 in bar down to 3.509 x 0.206 in
    const BAR_W = 3.509, BAR_H = 0.206, SCALE = BAR_W / 3.016;
    [['65%', 2.53, 3.894], ['22%', 1.957, 4.244], ['85%', 1.024, 4.593]].forEach(([label, fw, y]) => {
        s.addShape('roundRect', { x: 5.507, y, w: BAR_W, h: BAR_H, rectRadius: BAR_H / 2, fill: { color: 'F1F1F1' }, line: NOLINE });
        s.addText(label, {
            x: 5.507, y, w: fw * SCALE, h: BAR_H, shape: 'roundRect', rectRadius: BAR_H / 2, fill: { color: ACC1 }, line: NOLINE,
            fontFace: FH, fontSize: 7.88, color: WHITE, valign: 'middle', margin: [1, 5, 1, 5],
            shadow: shadow({ color: ACC2, opacity: 0.3, blur: 26, offset: 0.12, angle: 135 }),
        });
    });
    s.addText(LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero', bodyStyle({ x: 5.454, y: 2.91, w: 3.389, h: 0.604 }));
    s.addText('Banking App Ensures Fast and Safe Transactions', titleStyle({ x: 5.454, y: 1.19, w: 4.213, h: 1.464 }));
    addChrome(s);
}

function slide14(s) {
    laptop(s, 4.5, 1.24, 4.62, 3.06);
    gradFill(s, { x: 0.72, y: 2.813, w: 6.174, h: 1.926, r: 0.198, c2: mix(ACC1, OR_LT, 0.85) });
    [['Our Vision', 1.126], ['Our Mission', 3.945]].forEach(([head, x]) => {
        s.addShape('roundRect', { x, y: 3.852, w: 2.281, h: 1.183, rectRadius: 0.146, fill: { color: WHITE }, line: NOLINE, shadow: softShadow() });
        s.addText('Lorem I\u00a0is simply dummy like text of the printing',
            { x: x + 0.289, y: 4.342, w: 1.752, h: 0.495, fontFace: FB, fontSize: 8.25, color: GRAY50, lineSpacingMultiple: 1.5, valign: 'top' });
        s.addText(head, cardHead({ x: x + 0.289, y: 4.108, w: 1.606, h: 0.252 }));
        s.addShape('ellipse', { x: x + 2.063, y: 4.225, w: 0.436, h: 0.436, fill: { color: ACC1 }, line: NOLINE, shadow: softShadow() });
        icon(s, 'chevron', x + 2.148, 4.242, 0.19, WHITE);
    });
    s.addText([{ text: 'Mobile Apps', options: { breakLine: true } }, { text: 'Service' }], titleStyle({ x: 1.039, y: 1.479, w: 2.753, h: 1.01 }));
    s.addText('MOCKUP SECTION', kickerStyle({ x: 1.054, y: 1.082, w: 1.8, h: 0.29 }));
    s.addText(LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet',
        bodyStyle({ x: 1.107, y: 3.124, w: 5.123, h: 0.431, color: WHITE }));
    addChrome(s);
}

function slide15(s) {
    imac(s, 0.985, 1.05, 4.468, 3.02);
    s.addText(LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies. ', bodyStyle({ x: 6.004, y: 2.749, w: 2.897, h: 0.604 }));
    s.addText('Application Progress', titleStyle({ x: 6.004, y: 1.594, w: 2.897, h: 1.01 }));
    s.addText('Try Now', {
        x: 6.076, y: 3.977, w: 1.27, h: 0.354, shape: 'roundRect', rectRadius: 0.177, fill: { color: mix(ACC1, OR_LT2, 0.5) }, line: NOLINE,
        fontFace: FB, fontSize: 10.5, bold: true, color: BLACK, align: 'center', valign: 'middle',
    });
    addChrome(s);
}

function slide16(s) {
    s.addShape('roundRect', { x: 1.481, y: 1.967, w: 2.363, h: 3.196, rectRadius: 0.35, fill: { color: WHITE }, line: NOLINE, shadow: shadow({ color: BLACK, opacity: 0.12, blur: 100, offset: 0, angle: 90 }) });
    [[3.328, 3.528, 4.378, mix(ACC1, OR_LT2, 0.45), 'person'], [6.23, 6.429, 7.362, ACC1, 'users']]
        .forEach(([xc, xo, tx, col, ic]) => {
            s.addShape('roundRect', { x: xc, y: 3.422, w: 2.62, h: 1.09, rectRadius: 0.189, fill: { color: WHITE }, line: NOLINE, shadow: cardShadow() });
            s.addShape('ellipse', { x: xo, y: 3.594, w: 0.745, h: 0.745, fill: { color: col }, line: NOLINE, shadow: cardShadow() });
            icon(s, ic, xo + 0.235, 3.825, 0.29, WHITE);
            s.addText('Lorem I\u00a0is simplysaa dummy like text.', { x: tx, y: 3.846, w: 1.331, h: 0.495, fontFace: FB, fontSize: 8.25, color: GRAY50, lineSpacingMultiple: 1.5, valign: 'top' });
            s.addText('Description', cardHead({ x: tx, y: 3.612, w: 1.606, h: 0.252 }));
        });
    s.addText(LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies, ', bodyStyle({ x: 4.707, y: 2.679, w: 4.091, h: 0.431 }));
    s.addText('Budgeting App Helps You Manage Your Money', titleStyle({ x: 4.68, y: 1.133, w: 4.168, h: 1.464 }));
    addLogo(s);
}

function slide17(s) {
    tablet(s, 5.3145, 0.9735, 4.964, 3.661);
    s.addText(LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero', bodyStyle({ x: 0.978, y: 3.519, w: 3.389, h: 0.604 }));
    s.addText([{ text: 'Best Application', options: { breakLine: true } }, { text: 'For You!' }], titleStyle({ x: 0.978, y: 2.363, w: 3.389, h: 1.01 }));
    s.addText('2024', {
        x: 1.02, y: 1.581, w: 1.065, h: 0.439, shape: 'roundRect', rectRadius: 0.079, fill: { color: ACC1 }, line: NOLINE,
        fontFace: FH, fontSize: 10.13, color: WHITE, align: 'center', valign: 'middle',
    });
    s.addShape('triangle', { x: 1.1878, y: 1.9959, w: 0.147, h: 0.108, rotate: 180, fill: { color: ACC1 }, line: NOLINE });
    addChrome(s);
}

function slide18(s) {
    s.addText([
        { text: LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet eros quis urna.', options: { breakLine: true } },
        { text: ' ', options: { breakLine: true } },
        { text: 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.', options: { bold: true } },
    ], bodyStyle({ x: 1.035, y: 2.318, w: 3.725, h: 0.948 }));
    s.addText('Mobile App User Experience', titleStyle({ x: 1.035, y: 1.162, w: 3.389, h: 1.01 }));
    [[3.612, 3.573, 3.453, 1.411, 3.119], [4.318, 4.279, 4.159, 1.405, 3.2]].forEach(([yd, yr, yt, xt, wt]) => {
        s.addShape('ellipse', { x: 1.154, y: yd, w: 0.114, h: 0.114, fill: { color: ACC1 }, line: NOLINE });
        s.addShape('ellipse', { x: 1.115, y: yr, w: 0.192, h: 0.192, fill: NOFILL, line: { color: ACC1, width: 1.5 } });
        s.addText('Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt',
            bodyStyle({ x: xt, y: yt, w: wt, h: 0.431, bold: true, color: BLACK }));
    });
    addNav(s);
}

function slide19(s) {
    tablet(s, 0.655, 1.404, 4.745, 3.499);
    s.addText([
        { text: LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet eros quis urna.', options: { breakLine: true } },
        { text: ' ', options: { breakLine: true } },
        { text: 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.', options: { bold: true } },
    ], bodyStyle({ x: 5.865, y: 2.558, w: 3.725, h: 0.948 }));
    s.addText('New Messaging App Promises', titleStyle({ x: 5.865, y: 1.402, w: 3.899, h: 1.01 }));
    s.addShape('ellipse', { x: 5.984, y: 3.852, w: 0.114, h: 0.114, fill: { color: ACC1 }, line: NOLINE });
    s.addShape('ellipse', { x: 5.945, y: 3.813, w: 0.192, h: 0.192, fill: NOFILL, line: { color: ACC1, width: 1.5 } });
    s.addText('Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt',
        bodyStyle({ x: 6.241, y: 3.693, w: 3.119, h: 0.431, bold: true, color: BLACK }));
    s.addText('Try Now', {
        x: 5.971, y: 4.412, w: 1.27, h: 0.354, shape: 'roundRect', rectRadius: 0.177, fill: { color: mix(ACC1, OR_LT2, 0.5) }, line: NOLINE,
        fontFace: FB, fontSize: 10.5, bold: true, color: BLACK, align: 'center', valign: 'middle',
    });
    addChrome(s);
}

function slide20(s) {
    gradFill(s, { x: 3.37, y: 2.977, w: 5.63, h: 1.595, r: 0.164, c1: mix(ACC1, OR_LT, 0.05), c2: mix(ACC1, OR_LT, 0.9) });
    phone(s, 1.746, 1.03, 2.213, 4.396);
    s.addShape('roundRect', {
        x: 0.666, y: 3.168, w: 2.147, h: 1.37, rectRadius: 0.14, fill: { color: WHITE }, line: NOLINE,
        shadow: shadow({ color: BLACK, opacity: 0.1, blur: 80, offset: 0.6, angle: 50 }),
    });
    s.addText([{ text: '65', options: { fontSize: 24 } }, { text: '%', options: { fontSize: 18 } }],
        { x: 0.82, y: 3.375, w: 0.773, h: 0.464, fontFace: FH, bold: true, charSpacing: -1.35, color: ACC1, valign: 'middle' });
    s.addText('Awesome display', { x: 1.535, y: 3.439, w: 1.136, h: 0.396, fontFace: FH, fontSize: 12, bold: true, color: '0C3740', lineSpacingMultiple: 0.8, valign: 'middle' });
    s.addText('Lorem ipsum dolor sit amet consectetuer.', { x: 0.82, y: 3.868, w: 1.761, h: 0.418, fontFace: FH, fontSize: 9, color: GRAY50, lineSpacingMultiple: 1.2, valign: 'top' });
    s.addText('App Focuses on Sustainable', titleStyle({ x: 5.22, y: 1.632, w: 3.571, h: 1.01 }));
    s.addText('MOCKUP SECTION', kickerStyle({ x: 5.236, y: 1.234, w: 1.8, h: 0.29 }));
    s.addText('Service For Company', labelStyle({ x: 5.545, y: 3.223, w: 1.749, h: 0.29, charSpacing: 0, color: WHITE }));
    s.addText(LOREM_A + 'Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo.',
        bodyStyle({ x: 5.545, y: 3.512, w: 2.85, h: 0.776, color: WHITE }));
    icon(s, 'check', 5.24, 3.26, 0.25, WHITE);
    addChrome(s);
}

/** White footer strip with three chips — shared by the break and closing slides. */
function footerBar(s) {
    s.addShape('round2SameRect', { x: 0.536, y: 5.053, w: 9.016, h: 0.589, fill: { color: WHITE }, line: NOLINE, shadow: cardShadow() });
    [['Mobile Apps', 1.036], ['Modern & Clean', 4.532], ['BIApps', 7.95]]
        .forEach(([t, x]) => s.addText(t, { x, y: 5.24, w: 1.7, h: 0.252, fontFace: FH, fontSize: 10.5, color: BLACK, valign: 'top' }));
}

function slide21(s) {
    gradFill(s, { x: 0, y: 0, w: 10, h: 5.625 });
    s.addText([{ text: 'Break', options: { bold: true } }, { text: ' Slide', options: { bold: false, fontFace: 'Poppins Light' } }],
        { x: 1.378, y: 2.049, w: 7.04, h: 1.553, fontFace: FH, fontSize: 86.25, color: WHITE, valign: 'top' });
    s.addText([
        { text: 'sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.', options: { breakLine: true } },
        { text: 'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.', options: { bold: true } },
    ], bodyStyle({ x: 2.09, y: 3.629, w: 5.82, h: 0.431, color: WHITE, align: 'center' }));
    s.addText('IT\u2019S TIME TO', kickerStyle({ x: 4.1, y: 1.759, w: 1.8, h: 0.29, color: WHITE, align: 'center' }));
    addNav(s, { color: WHITE });
    addLogo(s, { color: WHITE });
    footerBar(s);
}

function slide22(s) {
    s.addText('Savings App Offers High-Interest', titleStyle({ x: 5.587, y: 1.371, w: 3.593, h: 0.909, fontSize: 24 }));
    s.addText('SERVICE SECTION', kickerStyle({ x: 5.603, y: 1.063, w: 1.647, h: 0.29, bold: false }));
    s.addText('Lorem ipsum dolor sit amet, adipiscing dolor sit amet, adipiscing elit. Mae cenas porttitor',
        bodyStyle({ x: 5.603, y: 2.403, w: 3.601, h: 0.431 }));
    ['01', '02', '03', '04'].forEach((n, i) => {
        const x = [0.84, 3.034, 5.228, 7.422][i];
        s.addText('Lorem ipsum dolor sit amet, adipiscing elit. Mae cenas porttitor', bodyStyle({ x, y: 4.183, w: 1.617, h: 0.604 }));
        s.addText('About Service', { x, y: 3.862, w: 1.218, h: 0.307, fontFace: FH, fontSize: 9, bold: true, color: DARK, lineSpacingMultiple: 1.5, valign: 'top' });
        s.addText(n, numStyle({ x, y: 3.406, w: 0.629, h: 0.505, fontSize: 24 }));
    });
    addLogo(s);
}

const PRICING = [
    { x: 0.884, y: 2.292, w: 2.169, h: 2.538, name: 'REGULAR APPS', nameSize: 10.5, price: 'FREE', items: 3, fs: 7.88, hdrY: 2.468, hdrW: 1.79, hdrH: 0.385, itemY: 3.022, itemDy: 0.2485, priceY: 3.814, btnY: 4.288, btnW: 1.24, btnH: 0.306 },
    { x: 3.645, y: 2.334, w: 2.169, h: 2.538, name: 'PRO APPS', nameSize: 10.5, price: '$ 100.00', items: 3, fs: 7.88, hdrY: 2.51, hdrW: 1.79, hdrH: 0.385, itemY: 3.064, itemDy: 0.2485, priceY: 3.854, btnY: 4.331, btnW: 1.24, btnH: 0.306 },
    { x: 6.405, y: 1.639, w: 2.718, h: 3.181, name: 'PREMIUM APPS', nameSize: 12, price: '$ 350.00', items: 4, fs: 8.25, hdrY: 1.848, hdrW: 2.284, hdrH: 0.492, itemY: 2.524, itemDy: 0.2775, priceY: 3.686, btnY: 4.216, btnW: 1.384, btnH: 0.341 },
];

function slide23(s) {
    s.addText('Pricing Table', titleStyle({ x: 0.806, y: 1.465, w: 3.113, h: 0.555 }));
    s.addText('Pricing Section', kickerStyle({ x: 0.822, y: 1.168, w: 1.275, h: 0.29, bold: false }));
    PRICING.forEach(c => {
        gradFill(s, { x: c.x, y: c.y, w: c.w, h: c.h, r: 0.175, c1: mix(ACC1, OR_LT, 0.05), c2: mix(ACC1, OR_LT, 0.9) });
        s.addShape('round2SameRect', { x: c.x + (c.w - c.hdrW) / 2, y: c.hdrY, w: c.hdrW, h: c.hdrH, fill: { color: WHITE, transparency: 85 }, line: NOLINE });
        s.addText(c.name, { x: c.x + 0.25, y: c.hdrY + 0.066, w: c.w - 0.5, h: 0.3, fontFace: FH, fontSize: c.nameSize, bold: true, color: WHITE, align: 'center', valign: 'top' });
        for (let i = 0; i < c.items; i++) {
            const y = c.itemY + i * c.itemDy;
            s.addText('Our Service Class 00' + (i + 1), { x: c.x + 0.471, y, w: c.w - 0.756, h: 0.24, fontFace: FH, fontSize: c.fs, color: WHITE, valign: 'top' });
            s.addShape('ellipse', { x: c.x + 0.391, y: y + 0.071, w: 0.07, h: 0.07, fill: { color: WHITE }, line: NOLINE });
        }
        s.addText(c.price, { x: c.x + 0.279, y: c.priceY, w: c.w - 0.558, h: 0.353, fontFace: FH, fontSize: 15, bold: true, color: WHITE, align: 'center', valign: 'top' });
        s.addText('Install Now', {
            x: c.x + (c.w - c.btnW) / 2, y: c.btnY, w: c.btnW, h: c.btnH, shape: 'roundRect', rectRadius: c.btnH / 2,
            fill: { color: WHITE }, line: NOLINE, fontFace: FH, fontSize: 7.88, bold: true, color: ACC1, align: 'center', valign: 'middle',
            shadow: shadow({ color: ACC1, opacity: 0.25, blur: 30, offset: 0.2, angle: 45 }),
        });
    });
    addChrome(s);
}

/* ---- SWOT family (24-27): giant letter panel + numbered items ---- */
function swotPanel(s, right, letter) {
    const x = right ? 6.014 : 0, r = 0.37;
    gradFill(s, { x, y: 1.014, w: 3.986, h: 3.597, r: right ? { tl: r, bl: r } : { tr: r, br: r }, dir: right ? 45 : 225 });
    s.addText(letter, {
        x: right ? 6.311 : 1.197, y: 1.534, w: 2.492, h: 2.613,
        fontFace: FH, fontSize: 149.25, bold: true, color: WHITE, align: 'center', valign: 'top',
    });
}

function swotHead(s, right, title, titleW) {
    const x = right ? 5.587 : 0.806;
    s.addText(title, titleStyle({ x, y: 1.276, w: titleW, h: 0.555 }));
    s.addText('SWOT Section', kickerStyle({ x: x + 0.016, y: 0.968, w: 1.275, h: 0.29, bold: false }));
}

function swotBadge(s, x, y, n) {
    s.addShape('ellipse', { x, y, w: 0.377, h: 0.377, fill: { color: mix(ACC1, OR_LT2, 0.5) }, line: NOLINE });
    s.addText(String(n), { x, y, w: 0.377, h: 0.377, fontFace: FB, fontSize: 10.13, color: WHITE, align: 'center', valign: 'middle' });
}

function slide24(s) {
    swotPanel(s, true, 'S');
    swotHead(s, false, 'Our Strengths', 3.272);
    [[1.049, 2.284, false], [3.112, 2.284, true], [1.049, 3.757, true], [3.112, 3.757, false]]
        .forEach(([x, y, orange], i) => {
            if (orange) gradFill(s, { x, y, w: 1.574, h: 0.938, r: 0.081 });
            else s.addShape('roundRect', { x, y, w: 1.574, h: 0.938, rectRadius: 0.081, fill: { color: WHITE }, line: NOLINE, shadow: cardShadow() });
            swotBadge(s, x - 0.155, y - 0.14, i + 1);
            s.addText('Lorem ipsum dolor amet, adipiscing ip sum dolor ',
                bodyStyle({ x: x + 0.216, y: y + 0.113, w: 1.245, h: 0.604, color: orange ? WHITE : GRAY50 }));
        });
    addChrome(s);
}

function slide25(s) {
    swotPanel(s, false, 'W');
    swotHead(s, true, 'Our Weaknesses', 3.928);
    [2.09, 2.792, 3.494, 4.196].forEach((y, i) => {
        swotBadge(s, 5.603, y, i + 1);
        s.addText('Lorem ipsum dolor amet, adipiscing dolor dolor amet, adipiscing ipsum dolor ',
            bodyStyle({ x: 6.152, y: y - 0.074, w: 3.035, h: 0.431 }));
    });
    addChrome(s);
}

function slide26(s) {
    swotPanel(s, true, 'O');
    swotHead(s, false, 'Our Opportunities', 3.928);
    [[1.386, 2.242, 0.822], [3.475, 2.242, 2.911], [1.386, 3.681, 0.822], [3.475, 3.681, 2.911]]
        .forEach(([x, y, xn], i) => {
            s.addText('Lorem ipsum dolor sit amet dolor sit amet dolor sit', bodyStyle({ x, y: y + 0.266, w: 1.218, h: 0.604 }));
            s.addText('Opportunities', { x, y, w: 1.218, h: 0.307, fontFace: FH, fontSize: 9, bold: true, color: DARK, lineSpacingMultiple: 1.5, valign: 'top' });
            s.addText('0' + (i + 1), numStyle({ x: xn, y, w: 0.629, h: 0.454 }));
        });
    addChrome(s);
}

function slide27(s) {
    swotPanel(s, false, 'T');
    swotHead(s, true, 'Our Threats', 2.647);
    [[2.227, 2.196], [3.657, 3.625]].forEach(([y, yn], i) => {
        s.addText(LOREM_A + 'Fusce posuere, magna magna sed', bodyStyle({ x: 6.066, y: y + 0.266, w: 3.091, h: 0.604 }));
        s.addText('Threats', { x: 6.066, y, w: 1.218, h: 0.307, fontFace: FH, fontSize: 9, bold: true, color: DARK, lineSpacingMultiple: 1.5, valign: 'top' });
        s.addText('0' + (i + 1), numStyle({ x: 5.587, y: yn, w: 0.629, h: 0.454 }));
    });
    addChrome(s);
}

function slide28(s) {
    // four overlapping step bubbles, painted back-to-front so 01 ends on top
    [[6.926, 7.066, '04', false], [4.892, 5.032, '03', true], [2.858, 2.998, '02', false], [0.824, 0.964, '01', true]]
        .forEach(([x, xn, n, orange]) => {
            if (orange) gradCircle(s, { x, y: 2.585, w: 2.266, h: 2.266 });
            else s.addShape('ellipse', { x, y: 2.585, w: 2.266, h: 2.266, fill: { color: WHITE }, line: NOLINE, shadow: cardShadow() });
            s.addShape('ellipse', { x: xn, y: 2.49, w: 0.648, h: 0.648, fill: { color: orange ? mix(ACC1, OR_LT, 0.3) : mix(ACC1, OR_LT, 0.5) }, line: NOLINE });
            s.addText(n, { x: xn, y: 2.49, w: 0.648, h: 0.648, fontFace: FB, fontSize: 10.13, color: WHITE, align: 'center', valign: 'middle' });
        });
    s.addText('Step For Install Apps', titleStyle({ x: 2.302, y: 1.531, w: 5.412, h: 0.555, align: 'center' }));
    s.addText('INFOGRAPHIC SECTION', kickerStyle({ x: 3.981, y: 1.223, w: 2.055, h: 0.29, align: 'center' }));
    [[1.247, 1.348, true], [3.281, 3.382, false], [5.315, 5.416, true], [7.349, 7.45, false]]
        .forEach(([xb, xt, onOrange]) => {
            s.addText('Lorem ipsum dolor sit amet, adipiscing elit. cenas porttitor', {
                x: xb, y: 3.528, w: 1.42, h: onOrange ? 0.706 : 0.604, align: 'center', valign: 'top',
                fontFace: onOrange ? FH : FB, fontSize: onOrange ? 8.25 : 7.88,
                color: onOrange ? WHITE : GRAY50, lineSpacingMultiple: onOrange ? 1.5 : 1.3,
            });
            s.addText('Tittle Here', {
                x: xt, y: 3.207, w: 1.218, h: 0.307, align: 'center', valign: 'top',
                fontFace: FH, fontSize: 9, bold: true, color: onOrange ? WHITE : DARK, lineSpacingMultiple: 1.5,
            });
        });
    addChrome(s);
}

const CONTACT = [
    { label: 'My Phone', value: '123-456-7890', ly: 0.753, vy: 1.108, iy: 1.071, isz: 0.286, kind: 'phone' },
    { label: 'My E-mail', value: 'hello@reallygreatsite.com', ly: 1.916, vy: 2.289, iy: 2.258, isz: 0.315, kind: 'mail' },
    { label: 'My Website', value: 'www.reallygreatsite.com', ly: 3.112, vy: 3.45, iy: 3.452, isz: 0.315, kind: 'web' },
    { label: 'My Adress', value: '123 Anywhere St., Any City', ly: 4.299, vy: 4.641, iy: 4.61, isz: 0.352, kind: 'pin' },
];

function slide29(s) {
    gradFill(s, { x: 0, y: 0, w: 5, h: 5.625 });
    CONTACT.forEach(c => {
        s.addText(c.value, { x: 6.85, y: c.vy, w: 2.4, h: 0.311, fontFace: FB, fontSize: 10.5, color: GRAY50, lineSpacingMultiple: 1.3, valign: 'top' });
        s.addText(c.label, { x: 6.39, y: c.ly, w: 1.438, h: 0.278, fontFace: FH, fontSize: 12, bold: true, color: '22301A', valign: 'top' });
        icon(s, c.kind, 6.42, c.iy, c.isz, 'A89490', WHITE);
    });
    s.addText('My Contact', { x: 0.731, y: 0.792, w: 4.033, h: 0.555, fontFace: FH, fontSize: 27, bold: true, color: WHITE, valign: 'top' });
    s.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim',
        bodyStyle({ x: 0.731, y: 1.498, w: 3.228, h: 0.604, color: WHITE }));
}

function slide30(s) {
    gradFill(s, { x: 0, y: 0, w: 10, h: 5.625 });
    phone(s, 1.457, 1.232, 2.64, 5.244);
    footerBar(s);
    s.addText('Thank You For Your Attention', {
        x: 4.75, y: 1.592, w: 4.181, h: 2.075, fontFace: FH, fontSize: 49.5, bold: true, color: WHITE, lineSpacingMultiple: 0.8, valign: 'top',
    });
    s.addText('Mobile Apps Presentation', {
        x: 4.75, y: 3.812, w: 4.26, h: 0.512, fontFace: FH, fontSize: 18, bold: true, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top',
    });
    addNav(s, { color: WHITE });
    addLogo(s, { color: WHITE });
    appScreen(s, {
        logo: { x: 2.4295, y: 2.229, w: 0.664, h: 0.701 },
        word: { x: 1.855, y: 3.2175, w: 1.8135, h: 0.4545, fontSize: 21 },
        btn: { x: 2.2075, y: 4.473, w: 1.1085, h: 0.309, fontSize: 10.5 },
    });
}

/* ================================================================== *
 * 5. Build
 * ================================================================== */
const BUILDERS = [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
    slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30,
];

function build() {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'BIAPPS', width: 10, height: 5.625 });
    pptx.layout = 'BIAPPS';
    pptx.theme = { headFontFace: FH, bodyFontFace: FB };
    pptx.title = 'Mobile Apps Presentation';

    BUILDERS.forEach(fn => {
        const slide = pptx.addSlide();
        slide.background = { color: WHITE };
        fn(slide);
    });

    return pptx.writeFile({ fileName: path.join(__dirname, '0858aacf-d6e1-4267-801e-5f738a1fa962_grok_final.pptx') });
}

build().then(f => console.log('wrote ' + f)).catch(e => { console.error(e); process.exit(1); });
