/**
 * "Doctor Profile" presentation template - rebuilt with pptxgenjs.
 *
 * Slide size 10 x 5.625in (16:9). Raster photos in the source deck are stock
 * placeholder art, so every one of them is redrawn here as a light grey panel
 * with a small "image" glyph (see imageBox()).
 *
 * Run: node 0a3dc204-01bd-427a-972a-1654e1bef891_grok_final.js
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */

const NAVY = '001834'; // theme accent1 - brand dark blue
const MINT = '8CDCC8'; // theme accent2 - brand highlight
const WHITE = 'FFFFFF';
const INK = '001227'; // slightly warmer navy used by one headline
const BODY = '3F3F3F'; // default paragraph grey
const MUTED = '7F7F7F';
const PLATE = 'F2F2F2'; // image placeholder fill
const GLYPH = '111111'; // image placeholder glyph
const STAR = 'FDBD26';
const GRID = 'BFBFBF';
const AXIS = 'D8D8D8';
const BLACK = '000000';

const HEAD = 'Lexend Medium'; // headings / numbers / labels
const SANS = 'DM Sans'; // body copy

// Google-Slides text insets: 0.075in sides, 0.0375in top/bottom -> points.
const INSET = [5.4, 5.4, 2.7, 2.7];

const LOREM =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do';
const LOREM_MED =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna aliqua. ';
const LOREM_LONG =
    'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
    'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis. ' +
    'Lorem ipsum dolor sit amet, consectetur adipiscing';

/* ------------------------------------------------------------------ *
 * Generic helpers
 * ------------------------------------------------------------------ */

/** Text box: top anchored, Google-Slides insets, no autofit. */
function text(slide, body, o) {
    slide.addText(body, Object.assign({ valign: 'top', margin: INSET, wrap: true }, o));
}

/** Headline (Lexend, 90% line spacing). */
function title(slide, body, o) {
    text(slide, body, Object.assign({ fontFace: HEAD, lineSpacingMultiple: 0.9 }, o));
}

/** 11pt body paragraph (DM Sans, 150% line spacing). */
function para(slide, body, o) {
    text(slide, body, Object.assign(
        { fontFace: SANS, fontSize: 11, color: BODY, lineSpacingMultiple: 1.5 }, o));
}

/** Straight rule. */
function rule(slide, x, y, w, color, opts) {
    slide.addShape('line', Object.assign(
        { x: x, y: y, w: w, h: 0, line: Object.assign({ color: color, width: 0.75 }, opts) }));
}

/**
 * Rectangle whose two TOP corners are rounded (PowerPoint "round2SameRect").
 * Combine with `rotate` to round any other pair of corners.
 */
function topRounded(slide, o) {
    const w = o.w, h = o.h, r = o.r, k = r * 0.5523; // 0.5523 = circular-arc bezier kappa
    slide.addShape('custGeom', {
        x: o.x, y: o.y, w: w, h: h, rotate: o.rotate || 0, fill: o.fill,
        points: [
            { x: 0, y: h },
            { x: 0, y: r },
            { x: r, y: 0, curve: { type: 'cubic', x1: 0, y1: r - k, x2: r - k, y2: 0 } },
            { x: w - r, y: 0 },
            { x: w, y: r, curve: { type: 'cubic', x1: w - r + k, y1: 0, x2: w, y2: r - k } },
            { x: w, y: h },
            { close: true },
        ],
    });
}

/* ------------------------------------------------------------------ *
 * Repeated deck furniture
 * ------------------------------------------------------------------ */

/**
 * Top-left brand lockup + optional "Page N" at the top right.
 * `light` picks the white-on-navy variant of the wordmark.
 */
function chrome(slide, opts) {
    const o = opts || {};
    const tint = o.light ? WHITE : NAVY;
    slide.addShape('roundRect', { x: 0.362, y: 0.206, w: 0.25, h: 0.25, fill: { color: MINT }, rectRadius: 0.0417 });
    // "caring hands cradling a heart" mark: outlined heart + cross, cupped by an arc
    slide.addShape('heart', {
        x: 0.432, y: 0.243, w: 0.11, h: 0.098, fill: { color: MINT }, line: { color: NAVY, width: 0.75 },
    });
    slide.addShape('mathPlus', { x: 0.469, y: 0.271, w: 0.037, h: 0.037, fill: { color: NAVY } });
    slide.addShape('blockArc', {
        x: 0.408, y: 0.262, w: 0.158, h: 0.158, fill: { color: NAVY },
        angleRange: [10, 170], arcThicknessRatio: 0.12,
    });
    text(slide, 'Doctor Profile', {
        x: 0.666, y: 0.209, w: 1.034, h: 0.25,
        fontFace: HEAD, fontSize: 9, color: tint, lineSpacingMultiple: 1.2,
    });
    if (o.page) {
        text(slide, 'Page ' + o.page, {
            x: 8.912, y: 0.21, w: 0.773, h: 0.227, align: 'right',
            fontFace: HEAD, fontSize: 9, color: o.pageLight ? WHITE : NAVY, lineSpacingMultiple: 1,
        });
    }
}

/** Mint "Get Started" pill (1.514 x 0.34in in the source deck). */
function pill(slide, x, y, label) {
    slide.addShape('roundRect', { x: x, y: y, w: 1.514, h: 0.34, fill: { color: MINT }, rectRadius: 0.17 });
    text(slide, label || 'Get Started', {
        x: x, y: y + 0.019, w: 1.514, h: 0.278, align: 'center',
        fontFace: HEAD, fontSize: 12, color: NAVY,
    });
}

/** "Your Text Here" kicker + paragraph, the deck's most reused text pairing. */
function kicker(slide, o) {
    const tint = o.color || NAVY;
    text(slide, o.head || 'Your Text Here', {
        x: o.x, y: o.y, w: o.w, h: 0.303, fontFace: HEAD, fontSize: 14, color: tint,
    });
    para(slide, o.body || LOREM, {
        x: o.x, y: o.y + 0.295, w: o.w, h: o.bodyH || 0.847, color: o.bodyColor || tint,
    });
}

/**
 * The "stacked photos" glyph that sits in the middle of every placeholder:
 * a back card peeking out bottom-left, a front frame, and inside it a sun
 * plus two mountains. `w` is the glyph width, height follows at 0.9 * w.
 */
function photoGlyph(slide, cx, cy, w) {
    const h = w * 0.9;
    const x = cx - w / 2, y = cy - h / 2;
    const fw = w * 0.82, fh = h * 0.76; // front frame
    const fx = x + w - fw, fy = y;
    const pad = w * 0.1; // frame border thickness
    slide.addShape('roundRect', { x: x, y: y + h - fh, w: fw, h: fh, fill: { color: GLYPH }, rectRadius: w * 0.1 });
    slide.addShape('roundRect', { x: fx, y: fy, w: fw, h: fh, fill: { color: GLYPH }, rectRadius: w * 0.1 });
    slide.addShape('rect', { x: fx + pad, y: fy + pad, w: fw - 2 * pad, h: fh - 2 * pad, fill: { color: PLATE } });
    slide.addShape('ellipse', {
        x: fx + pad * 1.5, y: fy + pad * 1.5, w: w * 0.18, h: w * 0.18, fill: { color: GLYPH },
    });
    slide.addShape('triangle', {
        x: fx + pad, y: fy + fh - pad - h * 0.3, w: fw * 0.45, h: h * 0.3, fill: { color: GLYPH },
    });
    slide.addShape('triangle', {
        x: fx + fw * 0.35, y: fy + fh - pad - h * 0.42, w: fw * 0.62, h: h * 0.42, fill: { color: GLYPH },
    });
}

/**
 * Stock-photo stand-in: rounded grey plate plus the photo glyph, sized the
 * way the original 16:9 artwork cover-scales inside its frame.
 */
function imageBox(slide, o) {
    // A roundRect with no explicit radius falls back to PowerPoint's 16.67%
    // default, so square-cornered plates must use a plain rect.
    slide.addShape(o.r ? 'roundRect' : 'rect', {
        x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate || 0,
        fill: { color: o.fill || PLATE }, rectRadius: o.r,
    });
    photoGlyph(slide, o.x + o.w / 2, o.y + o.h / 2, 0.98 * Math.max(o.w / 16, o.h / 9));
}

/**
 * Rounded stat card: kicker line, big figure, supporting sentence.
 * Used (with different palettes) on slides 3, 12 and 18.
 */
function statCard(slide, o) {
    slide.addShape('roundRect', {
        x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: o.fill }, rectRadius: o.r,
    });
    const tx = o.x + o.pad;
    text(slide, o.label, {
        x: tx, y: o.y + 0.113, w: o.labelW, h: 0.29,
        fontFace: HEAD, fontSize: 11, color: o.ink, lineSpacingMultiple: 1.3,
    });
    text(slide, o.value, {
        x: tx, y: o.y + 0.358, w: o.valueW, h: 0.682, fontFace: HEAD, fontSize: 36, color: o.ink,
    });
    para(slide, o.body || LOREM_SHORT, {
        x: tx, y: o.y + 0.988, w: o.bodyW, h: 0.582, color: o.bodyColor || o.ink,
    });
}

/** Five-star rating row (slide 15). */
function stars(slide, x, y) {
    for (let i = 0; i < 5; i++) {
        slide.addShape('star5', {
            x: x + i * 0.1948, y: y, w: 0.15, h: 0.143, fill: { color: STAR },
        });
    }
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

function slide01(pptx) {
    const s = pptx.addSlide();
    s.background = { color: NAVY };
    chrome(s, { light: true });
    title(s, 'Doctor Profile', { x: 0.482, y: 1.249, w: 7.863, h: 1.172, fontSize: 72, color: WHITE });
    title(s, 'Presentation Template', { x: 0.482, y: 2.421, w: 5.548, h: 0.487, fontSize: 27, color: WHITE });
    rule(s, 0.459, 3.931, 4.362, WHITE);
    kicker(s, { x: 5.179, y: 3.795, w: 3.961, color: WHITE });
}

function slide02(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 2 });
    title(s, 'Meet Dr. Jonathan', { x: 0.555, y: 0.784, w: 4.518, h: 1.893, fontSize: 60, color: NAVY });
    kicker(s, { x: 0.555, y: 3.164, w: 3.961, bodyColor: BODY });
    pill(s, 0.555, 4.54);
    imageBox(s, { x: 5.484, y: 0.615, w: 4.034, h: 4.56, r: 0.195 });
}

function slide03(pptx) {
    const s = pptx.addSlide();
    s.background = { color: NAVY };
    chrome(s, { light: true, page: 3, pageLight: true });
    title(s, 'Profile Overview', { x: 0.482, y: 0.694, w: 5.03, h: 1.898, fontSize: 60, color: WHITE });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
        'incididunt ut labore et dolore magna', { x: 0.482, y: 2.592, w: 4.828, h: 0.582, color: WHITE });

    statCard(s, {
        x: 6.238, y: 1.036, w: 3.036, h: 1.777, r: 0.213, fill: MINT, ink: NAVY,
        pad: 0.165, labelW: 1.751, valueW: 1.573, bodyW: 2.739,
        label: 'Internal Medicine', value: '98%',
    });
    statCard(s, {
        x: 6.238, y: 3.398, w: 3.036, h: 1.777, r: 0.213, fill: WHITE, ink: NAVY,
        pad: 0.165, labelW: 1.751, valueW: 1.573, bodyW: 2.739,
        label: 'Years of Experience',
        value: [{ text: '17', options: { fontSize: 36 } }, { text: 'Years', options: { fontSize: 18 } }],
        bodyColor: BODY,
    });
    imageBox(s, { x: 0.482, y: 3.398, w: 5.03, h: 1.777, r: 0.247 });
}

function slide04(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 4 });
    title(s, 'Vision & Mission', { x: 0.482, y: 0.78, w: 7.899, h: 1.075, fontSize: 66, color: NAVY });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
        'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
        'exercitation ullamco ', { x: 0.482, y: 1.794, w: 7.018, h: 0.582 });

    [
        { x: 0.482, fill: NAVY, ink: WHITE, head: 'Vision' },
        { x: 5.179, fill: MINT, ink: NAVY, head: 'Mission' },
    ].forEach(function (c) {
        s.addShape('roundRect', { x: c.x, y: 2.87, w: 4.256, h: 2.082, fill: { color: c.fill }, rectRadius: 0.202 });
        text(s, c.head, {
            x: c.x + 0.231, y: 3.074, w: 2.456, h: 0.372,
            fontFace: HEAD, fontSize: 15, color: c.ink, lineSpacingMultiple: 1.3,
        });
        para(s, LOREM_LONG, { x: c.x + 0.231, y: 3.519, w: 3.841, h: 1.112, color: c.ink });
    });
}

function slide05(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 5 });
    title(s, 'Medical Specializations', { x: 0.482, y: 0.654, w: 6.554, h: 1.893, fontSize: 60, color: INK });
    para(s, [
        { text: LOREM, options: { breakLine: true } },
        { text: '', options: { breakLine: true } },
        { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
            'incididunt ut labore et dolore' },
    ], { x: 5.357, y: 2.871, w: 4.162, h: 1.642 });
    pill(s, 5.357, 4.801);
    imageBox(s, { x: 0.482, y: 2.812, w: 4.518, h: 2.362, r: 0.215 });
}

function slide06(pptx) {
    const s = pptx.addSlide();
    s.background = { color: NAVY };
    // Banner photo and its 66%-opaque navy scrim share the same bottom rounding.
    topRounded(s, { x: 0, y: 0, w: 10, h: 3.036, r: 0.256, rotate: 180, fill: { color: PLATE } });
    photoGlyph(s, 5, 1.518, 0.98 * 10 / 16);
    topRounded(s, { x: 0, y: 0, w: 10, h: 3.036, r: 0.256, rotate: 180, fill: { color: NAVY, transparency: 34 } });
    chrome(s, { light: true, page: 6, pageLight: true });
    title(s, 'Hospital Affiliation', { x: 0.577, y: 0.657, w: 4.423, h: 1.893, fontSize: 60, color: WHITE });
    [0.577, 4.857].forEach(function (x) {
        text(s, 'Your Text Here', {
            x: x, y: 3.514, w: 2.456, h: 0.372,
            fontFace: HEAD, fontSize: 15, color: WHITE, lineSpacingMultiple: 1.3,
        });
        para(s, LOREM_LONG, { x: x, y: 3.959, w: 3.841, h: 1.112, color: WHITE });
    });
}

function slide07(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 7 });
    title(s, 'Services Offered', { x: 0.482, y: 0.674, w: 4.423, h: 1.893, fontSize: 60, color: NAVY });
    kicker(s, { x: 0.577, y: 3.221, w: 3.961, bodyColor: BODY });
    pill(s, 0.577, 4.676);

    s.addShape('roundRect', { x: 4.894, y: 0.844, w: 2.159, h: 4.173, fill: { color: MINT }, rectRadius: 0.259 });
    text(s, 'Preventive Care', {
        x: 5.116, y: 1.11, w: 1.856, h: 0.343, fontFace: HEAD, fontSize: 14, color: NAVY, lineSpacingMultiple: 1.3,
    });
    para(s, [
        { text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', options: { breakLine: true } },
        { text: '', options: { breakLine: true } },
        { text: 'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis. ' },
    ], { x: 5.135, y: 1.544, w: 1.753, h: 2.437, color: NAVY });
    text(s, '+58%', { x: 5.135, y: 4.1, w: 1.715, h: 0.682, fontFace: HEAD, fontSize: 36, color: NAVY });
    imageBox(s, { x: 7.36, y: 0.844, w: 2.159, h: 4.173, r: 0.228 });
}

function slide08(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    imageBox(s, { x: 0, y: 0, w: 10, h: 5.625, r: 0 });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 5.625, fill: { color: NAVY, transparency: 34 } });
    chrome(s, { light: true, page: 8, pageLight: true });
    title(s, 'Medical Overview', {
        x: 0.607, y: 0.681, w: 8.786, h: 0.985, fontSize: 60, color: WHITE, align: 'center',
    });

    s.addShape('roundRect', { x: 1.095, y: 2.529, w: 7.81, h: 2.645, fill: { color: WHITE }, rectRadius: 0.207 });
    [
        { x: 1.471, value: '280+' },
        { x: 4.077, value: '98%' },
        { x: 6.684, value: '36M' },
    ].forEach(function (col) {
        text(s, 'Your Text Here', {
            x: col.x, y: 2.804, w: 1.845, h: 0.313, align: 'center',
            fontFace: HEAD, fontSize: 12, color: NAVY, lineSpacingMultiple: 1.3,
        });
        text(s, col.value, {
            x: col.x, y: 3.102, w: 1.845, h: 0.682, align: 'center', fontFace: HEAD, fontSize: 36, color: NAVY,
        });
        para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor', {
            x: col.x, y: 3.769, w: 1.845, h: 1.112, align: 'center',
        });
    });
    [3.697, 6.303].forEach(function (x) {
        s.addShape('line', {
            x: x, y: 2.859, w: 0, h: 2.041, line: { color: NAVY, width: 0.75, transparency: 18 },
        });
    });
}

function slide09(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 9 });
    title(s, 'Education Background', {
        x: 0.884, y: 0.619, w: 8.233, h: 0.825, fontSize: 50, color: NAVY, align: 'center',
    });
    [0.482, 3.646, 6.81].forEach(function (x) {
        imageBox(s, { x: x, y: 1.818, w: 2.709, h: 2.193, r: 0.234 });
        text(s, 'Your Text Here', {
            x: x + 0.432, y: 4.291, w: 1.845, h: 0.313, align: 'center',
            fontFace: HEAD, fontSize: 12, color: NAVY, lineSpacingMultiple: 1.3,
        });
        para(s, LOREM_SHORT, { x: x, y: 4.604, w: 2.709, h: 0.582, align: 'center' });
    });
}

function slide10(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    // Navy panel: top-rounded rectangle turned a quarter turn counter-clockwise.
    topRounded(s, { x: 4.842, y: 0.467, w: 5.625, h: 4.69, r: 0.235, rotate: -90, fill: { color: NAVY } });
    chrome(s, { page: 10, pageLight: true });
    title(s, 'Certifications & Accreditations', {
        x: 0.482, y: 0.619, w: 4.518, h: 1.303, fontSize: 41, color: NAVY,
    });
    [
        { y: 0.868, labelW: 1.751, label: 'Advanced Cardiac Life Support (ACLS)', value: '2011', vy: 1.259, by: 1.014 },
        { y: 2.293, labelW: 1.751, label: 'Basic Life Support (BLS)', value: '2018', vy: 2.703, by: 2.458 },
        { y: 3.793, labelW: 1.537, label: 'Harvard Medical CME', value: '2023', vy: 4.148, by: 3.903 },
    ].forEach(function (row) {
        text(s, row.label, { x: 5.82, y: row.y, w: row.labelW, h: 0.429, fontFace: HEAD, fontSize: 11, color: WHITE });
        text(s, row.value, { x: 5.82, y: row.vy, w: 1.573, h: 0.682, fontFace: HEAD, fontSize: 36, color: WHITE });
        para(s, LOREM_SHORT, { x: 7.667, y: row.by, w: 1.959, h: 0.847, color: WHITE });
    });
    imageBox(s, { x: 0.482, y: 2.293, w: 4.209, h: 2.881, r: 0.215 });
}

function slide11(pptx) {
    const s = pptx.addSlide();
    s.background = { color: NAVY };
    chrome(s, { light: true, page: 11, pageLight: true });
    title(s, 'Clinical Experience', { x: 5.0, y: 0.817, w: 4.518, h: 1.712, fontSize: 54, color: WHITE });
    imageBox(s, { x: 0.482, y: 0.817, w: 4.09, h: 2.222, r: 0.197 });
    text(s, 'Internal Medicine Specialist', {
        x: 0.482, y: 3.285, w: 2.923, h: 0.303, fontFace: HEAD, fontSize: 14, color: WHITE,
    });
    text(s, '+340k', { x: 0.482, y: 3.595, w: 2.923, h: 0.833, fontFace: HEAD, fontSize: 45, color: WHITE });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
        'incididunt ut labore et dolore', { x: 0.482, y: 4.356, w: 4.09, h: 0.582, color: WHITE });
    kicker(s, { x: 5.089, y: 3.141, w: 3.961, color: WHITE });
    pill(s, 5.089, 4.596);
}

function slide12(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 12 });
    imageBox(s, { x: 5.833, y: 0.571, w: 4.167, h: 4.603, r: 0.216 });
    title(s, 'Professional Philosophy', { x: 0.482, y: 0.642, w: 4.959, h: 1.575, fontSize: 50, color: NAVY });
    para(s, LOREM_MED, { x: 0.482, y: 2.272, w: 4.518, h: 0.582 });
    statCard(s, {
        x: 0.482, y: 3.155, w: 2.833, h: 1.777, r: 0.213, fill: MINT, ink: NAVY,
        pad: 0.154, labelW: 2.557, valueW: 2.557, bodyW: 2.557, label: 'Your Text Here', value: '$45M',
    });
    statCard(s, {
        x: 3.679, y: 3.159, w: 2.833, h: 1.777, r: 0.213, fill: NAVY, ink: WHITE,
        pad: 0.154, labelW: 2.557, valueW: 2.557, bodyW: 2.557, label: 'Your Text Here', value: '+560K',
    });
}

function slide13(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 13 });
    title(s, 'What Our Clients Say', { x: 0.466, y: 0.789, w: 3.78, h: 1.303, fontSize: 41, color: NAVY });
    pill(s, 0.482, 2.593);
    text(s, '+76,8%', { x: 0.482, y: 3.213, w: 2.923, h: 0.833, fontFace: HEAD, fontSize: 45, color: NAVY });
    para(s, LOREM_MED, { x: 0.482, y: 3.99, w: 3.256, h: 0.847 });

    [
        { y: 1.01, name: 'Anita Clara', value: '87%', imgY: 0.864 },
        { y: 3.247, name: 'Josefin George', value: '56%', imgY: 3.102 },
    ].forEach(function (t) {
        text(s, t.name, { x: 5.237, y: t.y, w: 2.453, h: 0.278, fontFace: HEAD, fontSize: 12, color: NAVY });
        text(s, t.value, { x: 5.237, y: t.y + 0.252, w: 2.453, h: 0.682, fontFace: HEAD, fontSize: 36, color: NAVY });
        para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed', {
            x: 5.237, y: t.y + 0.873, w: 2.453, h: 0.582,
        });
        imageBox(s, { x: 7.773, y: t.imgY, w: 1.745, h: 1.746, r: 0.155 });
    });
}

function slide14(pptx) {
    const s = pptx.addSlide();
    s.background = { color: NAVY };
    chrome(s, { light: true, page: 14, pageLight: true });
    title(s, 'Insurance Accepted', {
        x: 0.669, y: 0.709, w: 8.661, h: 0.894, fontSize: 54, color: WHITE, align: 'center',
    });
    [
        { card: 0.94, inner: 1.217, fill: MINT },
        { card: 5.405, inner: 5.682, fill: WHITE },
    ].forEach(function (c) {
        topRounded(s, { x: c.card, y: 2.06, w: 3.655, h: 3.565, r: 0.189, fill: { color: c.fill } });
        imageBox(s, { x: c.inner, y: 2.357, w: 3.101, h: 1.818, r: 0.207 });
        para(s, LOREM_MED, { x: c.inner, y: 4.322, w: 3.256, h: 0.847, color: NAVY });
    });
}

function slide15(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    // Navy sidebar: top-rounded rectangle turned a quarter turn clockwise.
    topRounded(s, { x: -1.17, y: 1.157, w: 5.638, h: 3.298, r: 0.261, rotate: 90, fill: { color: NAVY } });
    chrome(s, { light: true, page: 15 });
    text(s, 'Team Profile', {
        x: 0.37, y: 0.609, w: 2.698, h: 0.53, align: 'center', fontFace: HEAD, fontSize: 27, color: WHITE,
    });
    imageBox(s, { x: 0.742, y: 1.512, w: 2.108, h: 2.249, r: 0.187 });
    text(s, 'Alex Paula', {
        x: 0.585, y: 3.982, w: 2.42, h: 0.379, align: 'center', fontFace: HEAD, fontSize: 18, color: WHITE,
    });
    text(s, 'General Doctor ', {
        x: 0.585, y: 4.339, w: 2.42, h: 0.252, align: 'center', fontFace: SANS, fontSize: 11, color: WHITE,
    });
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin', {
        x: 0.585, y: 4.614, w: 2.42, h: 0.701, align: 'center',
        fontFace: SANS, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.2,
    });

    [
        { x: 3.621, img: 3.795, star: 4.007, name: 'Sarah Brown', role: ' Internal Medicine' },
        { x: 5.746, img: 5.905, star: 6.133, name: 'Emily Smith', role: 'Pediatrics' },
        { x: 7.907, img: 8.014, star: 8.294, name: 'Avery Harrell', role: 'Cardiology ' },
    ].forEach(function (m) {
        imageBox(s, { x: m.img, y: 1.353, w: 1.504, h: 1.702, r: 0.167 });
        text(s, m.name, {
            x: m.x, y: 3.761, w: 1.703, h: 0.303, align: 'center', fontFace: HEAD, fontSize: 14, color: NAVY,
        });
        text(s, m.role, {
            x: m.x + 0.015, y: 3.982, w: 1.674, h: 0.29, align: 'center',
            fontFace: SANS, fontSize: 11, color: BODY, lineSpacingMultiple: 1.3,
        });
        stars(s, m.star, 4.342);
    });
}

function slide16(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 16 });
    title(s, 'Research & Publications', {
        x: 0.932, y: 0.638, w: 8.135, h: 0.757, fontSize: 45, color: NAVY, align: 'center',
    });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor ' +
        'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud ' +
        'exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.\u00a0', {
        x: 0.617, y: 1.49, w: 8.765, h: 0.582, align: 'center',
    });
    imageBox(s, { x: 0.482, y: 2.321, w: 9.036, h: 2.844, r: 0.216 });
}

function slide17(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 17 });
    title(s, 'Patient-Centered Approach', { x: 0.482, y: 0.714, w: 6.76, h: 1.575, fontSize: 50, color: NAVY });
    kicker(s, { x: 0.482, y: 2.86, w: 3.961, bodyColor: MUTED });
    pill(s, 0.482, 4.315);

    // Donut built from three pie wedges plus a white hub.
    [
        { x: 5.611, w: 3.305, from: 270.36, to: 42.55, fill: NAVY },
        { x: 5.617, w: 3.299, from: 193.44, to: 269.72, fill: PLATE },
        { x: 5.617, w: 3.299, from: 43.37, to: 192.48, fill: MINT },
    ].forEach(function (wedge) {
        s.addShape('pie', {
            x: wedge.x, y: 1.756, w: wedge.w, h: 3.279,
            angleRange: [wedge.from, wedge.to], fill: { color: wedge.fill },
        });
    });
    s.addShape('ellipse', { x: 6.425, y: 2.559, w: 1.682, h: 1.672, fill: { color: WHITE } });

    // Calendar glyph in the hub: two hanger tabs, navy body, hollow date area.
    [7.055, 7.365].forEach(function (x) {
        s.addShape('rect', { x: x, y: 3.086, w: 0.055, h: 0.115, fill: { color: NAVY } });
    });
    s.addShape('roundRect', { x: 6.933, y: 3.156, w: 0.617, h: 0.547, fill: { color: NAVY }, rectRadius: 0.07 });
    s.addShape('rect', { x: 7.02, y: 3.34, w: 0.443, h: 0.27, fill: { color: WHITE } });

    [
        { x: 6.189, y: 2.232, label: '15%', color: NAVY },
        { x: 8.121, y: 2.805, label: '45%', color: WHITE },
        { x: 6.291, y: 4.277, label: '50%', color: NAVY },
    ].forEach(function (l) {
        text(s, l.label, { x: l.x, y: l.y, w: 0.795, h: 0.379, fontFace: HEAD, fontSize: 18, color: l.color });
    });
}

function slide18(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 18 });
    title(s, 'Technology & Innovation', { x: 0.574, y: 0.532, w: 5.848, h: 1.439, fontSize: 45, color: NAVY });

    // Hand-drawn column chart: axis labels, grid lines, then the bars.
    const ticks = [
        { label: '60', y: 2.263, grid: 2.389, color: GRID },
        { label: '50', y: 2.697, grid: 2.823, color: GRID },
        { label: '40', y: 3.125, grid: 3.246, color: GRID },
        { label: '30', y: 3.547, grid: 3.673, color: GRID },
        { label: '20', y: 3.975, grid: 4.099, color: GRID },
        { label: '10', y: 4.4, grid: 4.522, color: GRID },
        { label: '0', y: 4.822, grid: 4.946, color: AXIS },
    ];
    ticks.forEach(function (t) {
        text(s, t.label, {
            x: 0.529, y: t.y, w: 0.304, h: 0.265, align: 'right',
            fontFace: SANS, fontSize: 9, color: BLACK, margin: [4.05, 4.05, 4.05, 4.05],
        });
        rule(s, 0.888, t.grid, 4.859, t.color);
    });
    [
        { x: 1.14, n: 1 }, { x: 2.355, n: 2 }, { x: 3.571, n: 3 }, { x: 4.784, n: 4 },
    ].forEach(function (c) {
        text(s, 'Category ' + c.n, {
            x: c.x, y: 4.955, w: 0.711, h: 0.257, align: 'center',
            fontFace: SANS, fontSize: 9, color: BLACK, lineSpacingMultiple: 1.3,
        });
    });
    [
        { x: 1.219, y: 3.665, w: 0.278, h: 1.282, fill: NAVY },
        { x: 1.536, y: 3.93, w: 0.276, h: 1.016, fill: MINT },
        { x: 2.435, y: 3.032, w: 0.278, h: 1.915, fill: NAVY },
        { x: 2.782, y: 3.833, w: 0.278, h: 1.114, fill: MINT },
        { x: 3.649, y: 4.275, w: 0.278, h: 0.672, fill: NAVY },
        { x: 3.995, y: 3.409, w: 0.278, h: 1.537, fill: MINT },
        { x: 4.863, y: 3.833, w: 0.278, h: 1.114, fill: NAVY },
        { x: 5.209, y: 4.478, w: 0.278, h: 0.468, fill: MINT },
    ].forEach(function (bar) {
        s.addShape('rect', { x: bar.x, y: bar.y, w: bar.w, h: bar.h, fill: { color: bar.fill } });
    });

    statCard(s, {
        x: 6.593, y: 1.252, w: 2.833, h: 1.777, r: 0.213, fill: MINT, ink: NAVY,
        pad: 0.154, labelW: 2.557, valueW: 2.557, bodyW: 2.557, label: 'Your Text Here', value: '$45M',
    });
    statCard(s, {
        x: 6.593, y: 3.398, w: 2.833, h: 1.777, r: 0.213, fill: NAVY, ink: WHITE,
        pad: 0.154, labelW: 2.557, valueW: 2.557, bodyW: 2.557, label: 'Your Text Here', value: '+560K',
    });
}

function slide19(pptx) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    chrome(s, { page: 19 });
    title(s, 'Product Prototype', { x: 0.482, y: 0.675, w: 7.628, h: 0.894, fontSize: 54, color: NAVY });
    // Tilted tablet mockup: placeholder screen with a thick dark bezel drawn on top.
    imageBox(s, { x: 5.572, y: 2.003, w: 3.68, h: 5.52, r: 0.196, rotate: -9 });
    s.addShape('roundRect', { // bezel: stroke centred on the mid-line of the frame
        x: 5.542, y: 1.958, w: 3.734, h: 5.585, rotate: -9.2, rectRadius: 0.292,
        line: { color: GLYPH, width: 10.2 },
    });
    pill(s, 0.482, 2.072);
    text(s, 'Internal Medicine Specialist', {
        x: 0.482, y: 2.67, w: 2.923, h: 0.303, fontFace: HEAD, fontSize: 14, color: NAVY,
    });
    text(s, '+230k', { x: 0.482, y: 2.994, w: 2.923, h: 0.833, fontFace: HEAD, fontSize: 45, color: NAVY });
    para(s, LOREM_MED, { x: 0.482, y: 3.771, w: 4.064, h: 0.847 });
}

function slide20(pptx) {
    const s = pptx.addSlide();
    s.background = { color: NAVY };
    chrome(s, { light: true });
    title(s, 'Thank You!', { x: 0.526, y: 1.084, w: 7.068, h: 1.388, fontSize: 86, color: WHITE });

    const contacts = [
        { badge: 0.526, textX: 0.929, textY: 2.741, w: 1.754, h: 0.451, icon: 'pin', label: '253 Mutton Town Road, South Dakota' },
        { badge: 2.929, textX: 3.37, textY: 2.823, w: 1.087, h: 0.255, icon: 'phone', label: '+1234567890' },
        { badge: 4.703, textX: 5.193, textY: 2.819, w: 1.831, h: 0.261, icon: 'mail', label: 'yourmail@company.com' },
    ];
    contacts.forEach(function (c) {
        const by = c.badge === 0.526 ? 2.762 : 2.766;
        s.addShape('roundRect', { x: c.badge, y: by, w: 0.401, h: 0.402, fill: { color: MINT }, rectRadius: 0.067 });
        contactIcon(s, c.icon, c.badge, by);
        text(s, c.label, {
            x: c.textX, y: c.textY, w: c.w, h: c.h,
            fontFace: HEAD, fontSize: 9, color: WHITE, lineSpacingMultiple: 1.3,
        });
    });
    rule(s, 0.526, 4.306, 3.932, WHITE);
    para(s, LOREM, { x: 5.179, y: 4.09, w: 3.961, h: 0.847, color: WHITE });
}

/** Small line-drawn navy pictograms inside the mint contact badges (slide 20). */
function contactIcon(slide, kind, bx, by) {
    const cx = bx + 0.2005, cy = by + 0.201;
    const stroke = { color: NAVY, width: 1 };
    if (kind === 'pin') {
        slide.addShape('teardrop', {
            x: cx - 0.076, y: cy - 0.112, w: 0.152, h: 0.152, rotate: 135,
            fill: { color: MINT }, line: stroke,
        });
        slide.addShape('ellipse', {
            x: cx - 0.032, y: cy - 0.088, w: 0.064, h: 0.064, fill: { color: MINT }, line: stroke,
        });
    } else if (kind === 'phone') {
        // Handset: the lower-left quadrant of a ring, capped by ear- and mouthpiece.
        slide.addShape('blockArc', {
            x: cx - 0.095, y: cy - 0.3, w: 0.38, h: 0.4,
            angleRange: [90, 180], arcThicknessRatio: 0.34, fill: { color: MINT }, line: stroke,
        });
        slide.addShape('roundRect', {
            x: cx - 0.107, y: cy - 0.118, w: 0.062, h: 0.08, rotate: -20,
            fill: { color: MINT }, line: stroke, rectRadius: 0.022,
        });
        slide.addShape('roundRect', {
            x: cx + 0.043, y: cy + 0.032, w: 0.08, h: 0.062, rotate: -20,
            fill: { color: MINT }, line: stroke, rectRadius: 0.022,
        });
    } else {
        slide.addShape('rect', {
            x: cx - 0.1, y: cy - 0.07, w: 0.2, h: 0.14, fill: { color: MINT }, line: stroke,
        });
        slide.addShape('custGeom', { // flap: a "V" from the two top corners
            x: cx - 0.1, y: cy - 0.07, w: 0.2, h: 0.14, fill: { color: MINT }, line: stroke,
            points: [{ x: 0, y: 0 }, { x: 0.1, y: 0.078 }, { x: 0.2, y: 0 }],
        });
    }
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

function build() {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'DECK_16x9', width: 10, height: 5.625 });
    pptx.layout = 'DECK_16x9';
    pptx.title = 'Doctor Profile';
    pptx.theme = { headFontFace: HEAD, bodyFontFace: SANS };

    [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
        slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
        .forEach(function (fn) { fn(pptx); });

    return pptx.writeFile({
        fileName: path.join(__dirname, '0a3dc204-01bd-427a-972a-1654e1bef891_grok_final.pptx'),
    });
}

build().then(function (f) { console.log('wrote ' + f); }, function (e) { console.error(e); process.exit(1); });
