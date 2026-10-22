/**
 * "Leadership Development" — 20-slide 16:9 deck (10 x 5.625 in) rebuilt with pptxgenjs.
 *
 * Run:  node 0351662e-920b-46f1-a35c-9a346d602e55_grok_final.js
 *       -> writes 0351662e-920b-46f1-a35c-9a346d602e55_grok_final.pptx beside this file.
 *
 * Photographs of the source deck are replaced by flat grey "[image]" placeholders
 * that keep the original position, size and clipping shape.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* =========================================================== design tokens */
const INK = '3F3F3F';   // headings + body copy
const DEEP = '3D524D';   // stat numbers
const LIME = 'C6CE85';   // subtitle accent
const WHITE = 'FFFFFF';
const PAGE_INK = '262626';   // "Page N" stamp
const PHOTO_BG = 'F2F2F2';   // placeholder tone for every bitmap
const PHOTO_LBL = 'A8ACAF';
const GOLD_DOT = 'D4B80B';

const SEMI = 'Poppins SemiBold';
const MED = 'Poppins Medium';
const POP = 'Poppins';
const BODY = 'Rubik';

/* Colour ramps, all running top-left -> bottom-right at 45 degrees. */
const BRAND = [[0, 'C6CE85'], [0.06, 'C6CE85'], [1, '829169']];   // cards, pills, panels
const COVER = [[0, 'C6CE85'], [0.72, '5F7059'], [1, '5C6D5C']];   // cover-slide circle
const BANNER = [[0, 'E4E8CD'], [1, 'B2BAB6']];                    // slanted translucent banner
const RING_GREEN = '93A170';   // thin decorative donuts
const BLOB_GREEN = 'CDD3C3';   // defocused background circles

/* 0.075in / 0.037in text insets of the original deck, expressed in points. */
const INSET = [5.4, 5.4, 2.7, 2.7];
// pptxgenjs rewrites the shadow object in place, so hand every shape a fresh one.
const cardShadow = () => ({ type: 'outer', angle: 45, blur: 8, offset: 3, color: '000000', opacity: 0.35 });

/* ================================================================ lorem copy */
const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ';
const L_EA = LOREM + 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea';
const L_CONSEQUAT = L_EA + ' commodo consequat.';
const L_QUIS_DOT = LOREM + 'Ut enim ad minim veniam, quis.';
const L_QUIS = LOREM + 'Ut enim ad minim veniam, quis';
const L_NISI = LOREM + 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut';
const L_LABORIS = LOREM + 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.';
const L_ALIQUIP = LOREM + 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip';
const L_EX = LOREM + 'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex';
const L_ULLAMCO = LOREM + 'Ut enim ad minim veniam, quis nostrud exercitation ullamco.';
const L_TEMPOR = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';
const L_ADIPISCING = 'Lorem ipsum dolor sit amet, consectetur adipiscing';
const L_SEDDO = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do';
const L_AMET = 'Lorem ipsum dolor sit amet';
const LABEL = 'Interactive Leadership';

/* ==================================================== colour ramp utilities */
function mix(a, b, t) {
    const ch = (hex, i) => parseInt(hex.substr(i * 2, 2), 16);
    const out = i => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t).toString(16).padStart(2, '0');
    return (out(0) + out(1) + out(2)).toUpperCase();
}

/** Colour of a multi-stop ramp at position t (0..1). */
function ramp(stops, t) {
    if (t <= stops[0][0]) return stops[0][1];
    for (let i = 1; i < stops.length; i++) {
        if (t <= stops[i][0]) {
            const [p0, c0] = stops[i - 1], [p1, c1] = stops[i];
            return mix(c0, c1, p1 === p0 ? 0 : (t - p0) / (p1 - p0));
        }
    }
    return stops[stops.length - 1][1];
}

/* =============================================== gradients built as polygons
 * pptxgenjs cannot emit <a:gradFill>, so every gradient is drawn as a fan of
 * flat-coloured custGeom slices cut at 45 degrees (or as plain rows/columns).
 */

/** Outline helpers: each returns an array of [x, y] points in inches. */
const rectPoly = (x, y, w, h) => [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];

function circlePoly(x, y, d, n = 56) {
    const r = d / 2, pts = [];
    for (let i = 0; i < n; i++) {
        const a = (2 * Math.PI * i) / n;
        pts.push([x + r + r * Math.cos(a), y + r + r * Math.sin(a)]);
    }
    return pts;
}

/** Rectangle whose left end is a semicircular cap (the deck's "stadium" panel). */
function stadiumPoly(x, y, w, h, n = 28) {
    const r = h / 2, pts = [[x + w, y]];
    for (let i = 0; i <= n; i++) {
        const a = -Math.PI / 2 - (Math.PI * i) / n;
        pts.push([x + r + r * Math.cos(a), y + r + r * Math.sin(a)]);
    }
    pts.push([x + w, y + h]);
    return pts;
}

/** Rectangle with the two corners on one side rounded (SWOT columns). */
function roundedEndPoly(x, y, w, h, r, atTop, n = 10) {
    const arc = (cx, cy, from) => {
        const out = [];
        for (let i = 0; i <= n; i++) {
            const a = from + (Math.PI / 2) * (i / n);
            out.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
        }
        return out;
    };
    return atTop
        ? [[x, y + h], [x, y + r], ...arc(x + r, y + r, Math.PI), ...arc(x + w - r, y + r, -Math.PI / 2), [x + w, y + h]]
        : [[x, y], [x + w, y], [x + w, y + h - r], ...arc(x + w - r, y + h - r, 0), ...arc(x + r, y + h - r, Math.PI / 2)];
}

/** Full-height banner that leans `lean` inches to the right as it rises. */
const leaningPoly = (x, y, w, h, lean) => [[x + lean, y], [x + w, y], [x + w - lean, y + h], [x, y + h]];

/** Cone seen slightly from above: wide flat top, rounded bottom (funnel stages). */
function conePoly(x, y, w, h, n = 20) {
    const inset = w * 0.16, rx = w / 2 - inset, ry = 0.1, cx = x + w / 2, cy = y + h - ry;
    const pts = [[x, y], [x + w, y]];
    for (let i = 0; i <= n; i++) pts.push([cx + rx * Math.cos((Math.PI * i) / n), cy + ry * Math.sin((Math.PI * i) / n)]);
    return pts;
}

/** Clip a convex polygon against the half-plane keep(pt) <= 0. */
function clipHalfPlane(poly, keep) {
    const out = [];
    for (let i = 0; i < poly.length; i++) {
        const a = poly[i], b = poly[(i + 1) % poly.length];
        const fa = keep(a), fb = keep(b);
        if (fa <= 0) out.push(a);
        if ((fa < 0 && fb > 0) || (fa > 0 && fb < 0)) {
            const t = fa / (fa - fb);
            out.push([a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t]);
        }
    }
    return out;
}

function addPolygon(slide, poly, color) {
    if (poly.length < 3) return;
    const xs = poly.map(q => q[0]), ys = poly.map(q => q[1]);
    const x = Math.min(...xs), y = Math.min(...ys);
    const w = Math.max(...xs) - x, h = Math.max(...ys) - y;
    if (w < 0.002 || h < 0.002) return;
    slide.addShape('custGeom', {
        x, y, w, h, fill: { color },
        points: poly.map(q => ({ x: q[0] - x, y: q[1] - y })).concat([{ close: true }]),
    });
}

/**
 * Fill a convex polygon with a ramp by slicing it into flat-coloured bands.
 * `axis` picks the direction: 'diag' (45 deg, the deck's default), 'y' or 'x'.
 */
function polyFill(slide, poly, stops, steps, axis) {
    const n = steps || 26, over = 0.006;
    const proj = axis === 'y' ? (q => q[1]) : axis === 'x' ? (q => q[0]) : (q => q[0] + q[1]);
    const vals = poly.map(proj);
    const lo = Math.min(...vals), hi = Math.max(...vals);
    for (let i = 0; i < n; i++) {
        const a = lo + ((hi - lo) * i) / n - over, b = lo + ((hi - lo) * (i + 1)) / n + over;
        let slice = clipHalfPlane(poly, pt => a - proj(pt));
        slice = clipHalfPlane(slice, pt => proj(pt) - b);
        addPolygon(slide, slice, ramp(stops, (i + 0.5) / n));
    }
}

const diagFill = (slide, poly, stops, steps) => polyFill(slide, poly, stops, steps, 'diag');

/** Brand-gradient rectangle, optionally sitting on the deck's soft card shadow. */
function gradCard(slide, o) {
    if (o.shadow) slide.addShape('rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: ramp(BRAND, 0.5) }, shadow: cardShadow() });
    diagFill(slide, rectPoly(o.x, o.y, o.w, o.h), o.stops || BRAND, o.steps);
}

/** Rounded pill with the brand ramp running left to right. */
function gradPill(slide, o) {
    const r = o.h / 2, mid = o.w - 2 * r, n = 12;
    slide.addShape('roundRect', { x: o.x, y: o.y, w: 2 * r, h: o.h, rectRadius: r, fill: { color: ramp(BRAND, r / o.w) } });
    slide.addShape('roundRect', { x: o.x + o.w - 2 * r, y: o.y, w: 2 * r, h: o.h, rectRadius: r, fill: { color: ramp(BRAND, (o.w - r) / o.w) } });
    for (let i = 0; i < n; i++) {
        slide.addShape('rect', {
            x: o.x + r + (mid * i) / n, y: o.y, w: mid / n + 0.01, h: o.h,
            fill: { color: ramp(BRAND, (r + (mid * (i + 0.5)) / n) / o.w) },
        });
    }
}

/* ============================================================ decorative art */
/** Thin open ring (the deck's "donut" accents). */
function ring(slide, o) {
    slide.addShape('donut', { x: o.x, y: o.y, w: o.d, h: o.d, rectRadius: o.t || 0.065, fill: { color: RING_GREEN } });
}

/** Defocused green circle: nested discs fading from white to green give the falloff. */
function blob(slide, o) {
    const n = 11;
    for (let i = 0; i < n; i++) {
        const k = 1 - (i / n) * 0.25;
        slide.addShape('ellipse', {
            x: o.x + (o.d * (1 - k)) / 2, y: o.y + (o.d * (1 - k)) / 2, w: o.d * k, h: o.d * k,
            fill: { color: mix(WHITE, BLOB_GREEN, (i + 1) / n) },
        });
    }
    slide.addShape('ellipse', { x: o.x + o.d * 0.13, y: o.y + o.d * 0.13, w: o.d * 0.74, h: o.d * 0.74, fill: { color: BLOB_GREEN } });
}

/* ================================================================ primitives */
/** Text box. Deck defaults: Rubik 9pt, ink grey, top-left, tight insets. */
function text(slide, str, o) {
    slide.addText(str, {
        x: o.x, y: o.y, w: o.w, h: o.h,
        fontFace: o.font || BODY, fontSize: o.size || 9, color: o.color || INK,
        bold: o.bold || false, align: o.align || 'left', valign: o.valign || 'top',
        lineSpacingMultiple: o.lh, margin: INSET, wrap: true,
    });
}

/** 36pt semibold slide heading. */
function heading(slide, str, o) {
    text(slide, str, Object.assign({ font: SEMI, size: 36, bold: true }, o));
}

/** "Page N" stamp inherited from the slide master (every slide but the cover). */
function pageNumber(slide, n) {
    text(slide, 'Page', { x: 8.944, y: 0.124, w: 0.546, h: 0.278, size: 12, color: PAGE_INK, align: 'center' });
    text(slide, String(n), { x: 9.49, y: 0.124, w: 0.385, h: 0.278, size: 12, color: PAGE_INK });
}

/** Grey stand-in for one of the deck's photographs. */
function photo(slide, o) {
    slide.addShape(o.shape || 'rect', Object.assign({
        x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: PHOTO_BG },
    }, o.shapeOpts || {}));
    text(slide, '[image]', {
        x: o.x, y: o.y + o.h / 2 - 0.16, w: o.w, h: 0.32, size: 10, color: PHOTO_LBL, align: 'center',
    });
}

/** Flat white card with the deck's soft drop shadow. */
function whiteCard(slide, o) {
    slide.addShape(o.shape || 'rect', Object.assign({
        x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: WHITE }, shadow: cardShadow(),
    }, o.shapeOpts || {}));
}

/* ====================================================== repeated composites */
/** Big number (optionally with a smaller unit) above a paragraph. */
function stat(slide, o) {
    const runs = o.unit
        ? [{ text: o.value, options: { fontSize: 27 } }, { text: o.unit, options: { fontSize: 21 } }]
        : o.value;
    slide.addText(runs, {
        x: o.x, y: o.y, w: o.vw || 1.5, h: 0.53, fontFace: MED, fontSize: 27,
        color: o.color || DEEP, align: o.align || 'left', valign: 'top', margin: INSET,
    });
    text(slide, o.body, {
        x: o.bx === undefined ? o.x : o.bx, y: o.y + 0.48, w: o.w, h: o.bh || 0.894,
        lh: 1.8, color: o.bodyColor || INK, align: o.align || 'left',
    });
}

/** Small 12pt label above a paragraph. */
function labeled(slide, o) {
    text(slide, o.label || LABEL, { x: o.x, y: o.y, w: 2.038, h: 0.278, font: MED, size: 12, color: o.color || INK });
    text(slide, o.body, { x: o.x, y: o.y + 0.259, w: o.w, h: o.bh || 0.621, lh: 1.8, color: o.color || INK });
}

/* ==================================================================== slides */
function slide01(pptx) {
    const s = pptx.addSlide();
    diagFill(s, circlePoly(-0.468, 4.071, 2.457), COVER, 24);
    blob(s, { x: 5.305, y: 0.2, d: 2.175 });
    blob(s, { x: 6.311, y: 1.665, d: 5.625 });
    [DEEP, LIME, 'F5DA37'].forEach((c, i) => // brand chips, cropped off the top edge
        s.addShape('rect', { x: 0.381 * i, y: -0.391, w: 0.327, h: 0.327, fill: { color: c } }));

    text(s, 'Leadership Development', { x: 0.246, y: 0.145, w: 2.535, h: 0.303, font: MED, size: 14 });
    text(s, '2025', { x: 9.17, y: 0.145, w: 0.612, h: 0.303, font: MED, size: 14 });
    text(s, 'Leadership Development', { x: 0.186, y: 0.736, w: 7.412, h: 2.297, font: SEMI, size: 66, bold: true });
    text(s, 'Presentation Template', { x: 0.227, y: 3.039, w: 6.548, h: 0.682, font: MED, size: 36, color: LIME });
    text(s, 'Business Presentation', { x: 2.402, y: 4.063, w: 2.312, h: 0.303, font: MED, size: 14 });
    text(s, L_EA, { x: 2.402, y: 4.298, w: 4.374, h: 0.894, lh: 1.8 });
}

function slide02(pptx) {
    const s = pptx.addSlide();
    ring(s, { x: -0.796, y: 2.942, d: 3.842 });
    blob(s, { x: 7.998, y: 3.4, d: 3.842 });
    photo(s, { x: 0.413, y: 0.8, w: 4.026, h: 4.026, shape: 'ellipse' });

    heading(s, 'Theories And Models', { x: 4.954, y: 1.019, w: 4.562, h: 1.287 });
    stat(s, { x: 4.989, y: 2.584, w: 4.338, value: '2025', body: L_EA });
    gradPill(s, { x: 5.0, y: 4.11, w: 1.433, h: 0.496 });
    text(s, 'Read More', { x: 5.228, y: 4.219, w: 0.977, h: 0.278, size: 12, color: WHITE });
    pageNumber(s, 2);
}

function slide03(pptx) {
    const s = pptx.addSlide();
    diagFill(s, leaningPoly(4.744, 0, 7.275, 5.625, 1.805), BANNER, 22);
    photo(s, { x: 5.525, y: 0.861, w: 3.43, h: 3.904 });

    heading(s, 'Skill Development', { x: 0.252, y: 0.861, w: 5.146, h: 0.682 });
    stat(s, { x: 0.273, y: 1.812, w: 4.338, value: '125k', body: L_EA });
    stat(s, { x: 0.273, y: 3.455, w: 4.338, value: '200k', body: L_EA });
    pageNumber(s, 3);
}

function slide04(pptx) {
    const s = pptx.addSlide();
    blob(s, { x: 0.116, y: 1.075, d: 2.702 });
    photo(s, { x: 1.57, y: 1.536, w: 3.741, h: 3.904 });

    heading(s, 'Case Studies', { x: 1.694, y: 0.612, w: 3.494, h: 0.682 });
    stat(s, { x: 5.475, y: 1.452, w: 4.338, bh: 1.166, value: '69', unit: '%', body: L_CONSEQUAT });
    stat(s, { x: 5.475, y: 3.858, w: 4.338, bh: 1.166, value: '79', unit: '%', body: L_CONSEQUAT });
    pageNumber(s, 4);
}

function slide05(pptx) {
    const s = pptx.addSlide();
    heading(s, 'Assessment And Tools', { x: 0.252, y: 0.525, w: 4.29, h: 1.321 });
    photo(s, { x: 6.073, y: 0.557, w: 3.741, h: 4.883 });

    gradCard(s, { x: 0.326, y: 2.44, w: 2.326, h: 3.0, shadow: true });
    text(s, '$20.00', { x: 0.771, y: 2.66, w: 1.436, h: 0.53, font: MED, size: 27, color: WHITE });
    text(s, L_QUIS_DOT, { x: 0.453, y: 3.824, w: 2.072, h: 1.439, lh: 1.8, color: WHITE, align: 'center' });

    whiteCard(s, { x: 3.199, y: 2.44, w: 2.326, h: 3.0 });
    text(s, '$20.00', { x: 3.644, y: 2.66, w: 1.436, h: 0.53, font: MED, size: 27 });
    text(s, L_QUIS_DOT, { x: 3.326, y: 3.824, w: 2.072, h: 1.439, lh: 1.8, align: 'center' });
    pageNumber(s, 5);
}

function slide06(pptx) {
    const s = pptx.addSlide();
    // layout art: stadium-shaped green panel, semicircular cap on the left
    diagFill(s, stadiumPoly(4.12, 1.159, 5.881, 3.306), BRAND, 26);
    photo(s, { x: 4.27, y: 1.29, w: 3.046, h: 3.046, shape: 'ellipse' });

    heading(s, 'Activities And Exercises', { x: 0.272, y: 1.228, w: 3.999, h: 1.287 });
    stat(s, { x: 0.272, y: 2.751, w: 3.711, bh: 1.166, value: '+450', body: L_EA });
    text(s, LABEL, { x: 7.642, y: 1.773, w: 2.038, h: 0.278, font: MED, size: 12, color: WHITE, align: 'center' });
    text(s, L_NISI, { x: 7.509, y: 2.14, w: 2.305, h: 1.712, lh: 1.8, color: WHITE, align: 'center' });
    pageNumber(s, 6);
}

function slide07(pptx) {
    const s = pptx.addSlide();
    heading(s, 'Action Plan', { x: 0.264, y: 0.557, w: 3.052, h: 0.682 });
    photo(s, { x: 3.792, y: 0.557, w: 3.741, h: 4.883 });
    photo(s, { x: 7.688, y: 0.557, w: 2.126, h: 4.883 });

    gradCard(s, { x: 0.305, y: 1.557, w: 4.343, h: 1.646, shadow: true });
    stat(s, { x: 0.507, y: 1.693, w: 3.941, value: '+500', body: L_LABORIS, color: WHITE, bodyColor: WHITE });

    whiteCard(s, { x: 0.305, y: 3.429, w: 4.343, h: 1.646 });
    stat(s, { x: 0.507, y: 3.565, w: 3.941, value: '+600', body: L_LABORIS, color: INK });
    pageNumber(s, 7);
}

function slide08(pptx) {
    const s = pptx.addSlide();
    blob(s, { x: 8.074, y: -1.515, d: 3.603 });
    photo(s, { x: 0.186, y: 1.76, w: 4.605, h: 3.68 });

    heading(s, 'Management Framework', { x: 0.231, y: 0.369, w: 5.854, h: 1.287 });
    stat(s, { x: 5.0, y: 1.644, w: 4.338, value: '200k', body: L_EA });
    stat(s, { x: 5.0, y: 4.067, w: 4.338, value: '300k', body: L_EA });
    pageNumber(s, 8);
}

function slide09(pptx) {
    const s = pptx.addSlide();
    photo(s, {
        x: -1.546, y: 0, w: 6.658, h: 5.625, shape: 'parallelogram',
        shapeOpts: { rectRadius: 1.406, flipH: true },
    });
    heading(s, 'Team Motivation ', { x: 5.0, y: 0.784, w: 5.639, h: 0.682 });
    stat(s, { x: 5.032, y: 1.705, w: 4.338, value: '97', unit: '%', body: L_EA });
    stat(s, { x: 5.032, y: 3.364, w: 4.338, value: '80', unit: '%', body: L_EA });
    pageNumber(s, 9);
}

function slide10(pptx) {
    const s = pptx.addSlide();
    blob(s, { x: 3.908, y: 3.824, d: 3.603 });
    photo(s, { x: 3.673, y: 0, w: 3.952, h: 4.883 });

    heading(s, 'Developing Strategic Thinking', { x: 0.202, y: 0.742, w: 4.345, h: 1.893 });
    stat(s, { x: 0.262, y: 3.354, w: 3.171, bh: 1.166, value: '300k', body: L_ALIQUIP });
    stat(s, { x: 7.714, y: 0.989, w: 2.113, bh: 1.166, value: '130k', body: LOREM });
    stat(s, { x: 7.714, y: 3.354, w: 2.113, bh: 1.166, value: '400k', body: LOREM });
    pageNumber(s, 10);
}

function slide11(pptx) {
    const s = pptx.addSlide();
    diagFill(s, leaningPoly(3.306, 0, 8.635, 5.625, 1.805), BANNER, 22);

    s.addShape('ellipse', { x: 2.908, y: 1.234, w: 3.157, h: 3.157, fill: { color: WHITE } });
    stat(s, {
        x: 4.052, y: 1.99, vw: 0.869, bx: 3.22, w: 2.533, bh: 1.166,
        value: '69', unit: '%', body: L_QUIS_DOT, align: 'center',
    });
    // the photo circle sits on top, cropping the left edge of the paragraph
    photo(s, { x: 0.413, y: 1.177, w: 3.272, h: 3.271, shape: 'ellipse' });

    heading(s, 'Measuring Leadership', { x: 6.44, y: 1.489, w: 2.968, h: 1.287 });
    labeled(s, { x: 6.471, y: 2.984, w: 3.224, bh: 0.894, body: L_QUIS });
    pageNumber(s, 11);
}

function slide12(pptx) {
    const s = pptx.addSlide();
    blob(s, { x: 1.88, y: -1.046, d: 3.603 });
    photo(s, { x: 0.186, y: 0.551, w: 2.906, h: 4.889 });
    photo(s, { x: 3.309, y: 1.485, w: 2.906, h: 3.955 });

    heading(s, 'Coaching Techniques', { x: 6.473, y: 1.326, w: 3.222, h: 1.287 });
    labeled(s, { x: 6.526, y: 2.73, w: 3.224, bh: 0.348, body: L_ADIPISCING });
    stat(s, { x: 6.494, y: 3.615, w: 2.801, bh: 1.439, value: '200k', body: L_EX });
    pageNumber(s, 12);
}

/** Progress-meter card, used twice on slide 13. */
function meterCard(s, o) {
    if (o.dark) gradCard(s, { x: 5.352, y: o.y, w: 4.343, h: 1.24, shadow: true });
    else whiteCard(s, { x: 5.352, y: o.y, w: 4.343, h: 1.24 });

    text(s, 'Your Value', { x: 5.524, y: o.y + 0.086, w: 1.279, h: 0.341, font: SEMI, size: 11, bold: true, color: o.ink, lh: 1.5 });
    text(s, '75%', { x: 8.242, y: o.y + 0.075, w: 0.74, h: 0.341, font: SEMI, size: 11, bold: true, color: o.ink, lh: 1.5, align: 'center' });
    s.addShape('rect', { x: 5.836, y: o.y + 0.419, w: 3.688, h: 0.063, fill: { color: o.track } });
    s.addShape('rect', { x: 5.59, y: o.y + 0.419, w: 3.031, h: 0.063, fill: { color: o.bar } });
    text(s, LOREM, { x: 5.524, y: o.y + 0.545, w: 3.789, h: 0.621, lh: 1.8, color: o.ink });
}

function slide13(pptx) {
    const s = pptx.addSlide();
    ring(s, { x: 6.078, y: -0.955, d: 2.68, t: 0.058 });
    heading(s, 'Measuring Leadership Impact', { x: 0.241, y: 0.981, w: 5.473, h: 1.287 });
    photo(s, { x: 0.186, y: 2.812, w: 4.814, h: 2.628 });

    meterCard(s, { y: 2.812, dark: true, ink: WHITE, track: 'FCF7D6', bar: 'F8E886' });
    meterCard(s, { y: 4.2, dark: false, ink: INK, track: 'AABFBA', bar: '3D524C' });
    pageNumber(s, 13);
}

function slide14(pptx) {
    const s = pptx.addSlide();
    blob(s, { x: 8.012, y: 3.639, d: 3.603 });
    photo(s, { x: 5.0, y: 0.49, w: 4.814, h: 2.323 });
    photo(s, { x: 0.186, y: 2.812, w: 2.906, h: 2.628 });

    heading(s, 'What Is Leadership?', { x: 0.448, y: 0.395, w: 3.893, h: 1.287 });
    labeled(s, { x: 0.448, y: 1.719, w: 4.084, body: LOREM });
    stat(s, { x: 3.415, y: 3.61, w: 2.807, bh: 1.439, value: '$50.00', body: L_ALIQUIP });
    stat(s, { x: 6.888, y: 3.61, w: 2.807, bh: 1.439, value: '$20.00', body: L_ALIQUIP });
    pageNumber(s, 14);
}

function slide15(pptx) {
    const s = pptx.addSlide();
    ring(s, { x: 8.318, y: -1.239, d: 2.68, t: 0.058 });
    photo(s, { x: 0.186, y: 0.185, w: 4.814, h: 5.255 });

    heading(s, 'Emotional Intelligence', { x: 5.306, y: 0.375, w: 3.56, h: 1.287 });
    labeled(s, { x: 5.306, y: 1.789, w: 4.084, body: LOREM });

    gradCard(s, { x: 5.355, y: 2.812, w: 2.072, h: 2.628, shadow: true });
    text(s, '+120k', { x: 5.778, y: 3.123, w: 1.227, h: 0.53, font: MED, size: 27, color: WHITE });
    text(s, L_TEMPOR, { x: 5.435, y: 4.297, w: 1.912, h: 0.894, lh: 1.8, color: WHITE, align: 'center' });

    whiteCard(s, { x: 7.742, y: 2.812, w: 2.072, h: 2.628 });
    text(s, '+120k', { x: 8.164, y: 3.123, w: 1.227, h: 0.53, font: MED, size: 27 });
    text(s, L_TEMPOR, { x: 7.822, y: 4.297, w: 1.912, h: 0.894, lh: 1.8, align: 'center' });
    pageNumber(s, 15);
}

/* ---- slide 16: SWOT ------------------------------------------------------ */
const SWOT = [
    { x: 1.092, letter: 'S', label: 'Strength', ink: '3D524C', top: '405650', bot: '7D9D95', roundTop: false },
    { x: 3.174, letter: 'W', label: 'Weakness', ink: 'C5CE85', top: 'C7D088', bot: '95A140', roundTop: true },
    { x: 5.256, letter: 'O', label: 'Opportunity', ink: '919864', top: '606442', bot: '929864', roundTop: false },
    { x: 7.339, letter: 'T', label: 'Threat', ink: '55726B', top: '597770', bot: '374B46', roundTop: true },
];
const COL_W = 1.819, COL_H = 2.277, COL_Y = 2.048, COL_R = 0.303;

/** One SWOT column: two corners rounded, filled with a vertical ramp. */
function swotColumn(s, col) {
    polyFill(s, roundedEndPoly(col.x, COL_Y, COL_W, COL_H, COL_R, col.roundTop),
        [[0, col.top], [1, col.bot]], 20, 'y');

    text(s, col.label, { x: col.x + 0.082, y: 2.642, w: 1.655, h: 0.328, font: MED, size: 15, color: WHITE, align: 'center' });
    [3.081, 3.565].forEach((y, i) => {
        s.addShape('ellipse', { x: col.x + 0.345, y: y + 0.086, w: 0.183, h: 0.183, fill: { color: WHITE, transparency: 100 }, line: { color: WHITE, width: 0.75 } });
        s.addText('\u2713', { x: col.x + 0.345, y: y + 0.086, w: 0.183, h: 0.183, fontFace: BODY, fontSize: 8, color: WHITE, align: 'center', valign: 'middle', margin: 0 });
        text(s, i === 0 ? L_AMET : L_AMET + '.', { x: col.x + 0.528, y, w: 1.15, h: 0.439, lh: 1.2, color: WHITE });
    });
    // white tile carrying the SWOT initial, overlapping the column top
    whiteCard(s, { x: col.x + 0.494, y: 1.632, w: 0.831, h: 0.831 });
    text(s, col.letter, { x: col.x + 0.539, y: 1.707, w: 0.723, h: 0.682, font: SEMI, size: 36, bold: true, color: col.ink, align: 'center' });
}

function slide16(pptx) {
    const s = pptx.addSlide();
    text(s, 'Marketing Target Analysis', { x: 1.586, y: 0.521, w: 7.078, h: 0.682, font: SEMI, size: 36, bold: true, align: 'center' });
    SWOT.forEach(col => swotColumn(s, col));
    text(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin convallis, arcu in aliquam molestie, nibh augue aliquam arcu, eu interdum sem nisl sit amet ligula. ',
        { x: 2.417, y: 4.571, w: 5.416, h: 0.47, lh: 1.3, align: 'center' });
    [[0.903, 5.342], [1.028, 5.467]].forEach(([x, y]) =>
        s.addShape('ellipse', { x, y, w: 0.102, h: 0.102, fill: { color: GOLD_DOT } }));
    pageNumber(s, 16);
}

/* ---- slide 17: funnel ---------------------------------------------------- */
const FUNNEL = [
    { y: 1.321, x: 6.026, w: 2.997, h: 0.911, rim: '8F8020', face: 'C0AB2B', label: 'Programming', lx: 6.749, lw: 1.68, badge: [5.921, 1.511] },
    { y: 2.323, x: 6.269, w: 2.51, h: 0.898, rim: '3F5550', face: '55726B', label: 'Networking', lx: 6.781, lw: 1.486, badge: [6.154, 2.543] },
    { y: 3.306, x: 6.467, w: 2.114, h: 0.885, rim: 'A8B548', face: 'C5CE85', label: 'Database', lx: 6.781, lw: 1.486, badge: [6.293, 3.52] },
];
const FUNNEL_STATS = [
    { x: 0.982, value: '198K', color: 'C0AB2B', vx: 1.138, vw: 1.222, name: 'Programming' },
    { x: 2.692, value: '216K', color: '55726B', vx: 2.872, vw: 1.176, name: 'Networking' },
    { x: 4.402, value: '302K', color: 'C5CE85', vx: 4.501, vw: 1.338, name: 'Database' },
];

function slide17(pptx) {
    const s = pptx.addSlide();
    text(s, 'Creative Funnel Diagram  ', { x: 0.967, y: 0.58, w: 4.355, h: 1.181, font: SEMI, size: 36, bold: true, lh: 0.9 });
    text(s, 'Data Description', { x: 0.98, y: 2.179, w: 2.395, h: 0.278, font: MED, size: 12 });
    text(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s. ",
        { x: 0.975, y: 2.475, w: 3.755, h: 0.651, lh: 1.3 });

    FUNNEL_STATS.forEach(c => {
        whiteCard(s, { x: c.x, y: 3.523, w: 1.535, h: 1.535 });
        text(s, c.value, { x: c.vx, y: 3.719, w: c.vw, h: 0.429, font: MED, size: 21, color: c.color, align: 'center' });
        text(s, c.name, { x: c.x + 0.043, y: 4.144, w: 1.45, h: 0.279, font: MED, size: 11, lh: 1.2, align: 'center' });
        text(s, 'Lorem ipsum dolor sit amet elit', { x: c.x + 0.147, y: 4.396, w: 1.241, h: 0.452, lh: 1.3, align: 'center' });
    });

    // dark rod skewered through the three stages, ending in a down arrow
    [[0.725, 0.705], [1.715, 0.705], [2.71, 0.705], [4.078, 0.5]].forEach(([y, h]) =>
        s.addShape('rect', { x: 7.409, y, w: 0.23, h, fill: { color: INK } }));
    s.addShape('triangle', { x: 7.168, y: 4.538, w: 0.713, h: 0.38, rotate: 180, fill: { color: INK } });

    FUNNEL.forEach(f => {
        diagFill(s, conePoly(f.x, f.y, f.w, f.h), [[0, mix(f.face, WHITE, 0.62)], [0.45, f.face], [1, f.face]], 14);
        s.addShape('ellipse', { x: f.x, y: f.y - 0.11, w: f.w, h: 0.22, fill: { color: f.rim } });
        text(s, f.label, { x: f.lx, y: f.y + 0.343, w: f.lw, h: 0.328, font: MED, size: 15, color: WHITE, align: 'center' });
        s.addShape('ellipse', { x: f.badge[0], y: f.badge[1], w: 0.443, h: 0.443, fill: { color: WHITE } });
        s.addShape('rect', { x: f.badge[0] + 0.111, y: f.badge[1] + 0.111, w: 0.22, h: 0.22, fill: { color: f.face } });
    });
    pageNumber(s, 17);
}

/* ---- slide 18: join our team --------------------------------------------- */
function featureCard(s, o) {
    if (o.dark) gradCard(s, { x: 0.645, y: o.y, w: 3.839, h: 1.08, shadow: true });
    else whiteCard(s, { x: 0.645, y: o.y, w: 3.839, h: 1.08 });

    if (o.dark) whiteCard(s, { x: 0.927, y: o.y + 0.196, w: 0.688, h: 0.688 });
    else gradCard(s, { x: 0.927, y: o.y + 0.196, w: 0.688, h: 0.688, steps: 10, shadow: true });
    // screen-cast glyph: two overlapping outlined screens
    s.addShape('rect', { x: 1.049, y: o.y + 0.318, w: 0.3, h: 0.22, fill: { color: WHITE, transparency: 100 }, line: { color: o.icon, width: 1 } });
    s.addShape('rect', { x: 1.192, y: o.y + 0.43, w: 0.3, h: 0.22, fill: { color: WHITE, transparency: 100 }, line: { color: o.icon, width: 1 } });

    text(s, o.title, { x: 1.752, y: o.y + 0.167, w: 2.375, h: 0.278, font: POP, size: 12, bold: true, color: o.ink });
    text(s, L_SEDDO, { x: 1.752, y: o.y + 0.459, w: 2.375, h: 0.452, lh: 1.3, color: o.ink });
}

function slide18(pptx) {
    const s = pptx.addSlide();
    // tablet mock-up (a group rotated 90 degrees in the source deck)
    s.addShape('roundRect', { x: 3.416, y: 0.46, w: 6.303, h: 4.648, rectRadius: 0.28, fill: { color: '111111' }, line: { color: 'DDDDDD', width: 2 } });

    text(s, 'Join Our\nTeam', { x: 0.58, y: 0.729, w: 2.375, h: 1.287, font: SEMI, size: 36, bold: true });
    featureCard(s, { y: 2.462, dark: true, ink: WHITE, icon: '3D524C', title: 'Feature 01' });
    featureCard(s, { y: 3.716, dark: false, ink: INK, icon: WHITE, title: 'Feature 03' });
    // the screen is the last shape in the source deck, so it clips both cards
    photo(s, { x: 3.57, y: 0.611, w: 6.003, h: 4.379 });
    pageNumber(s, 18);
}

/* ---- slide 19: contact --------------------------------------------------- */
const CONTACTS = [
    { y: 1.834, label: 'Address', dark: false },
    { y: 2.926, label: 'Phone', dark: true },
    { y: 4.018, label: 'Website', dark: false },
];

function slide19(pptx) {
    const s = pptx.addSlide();
    // laptop mock-up bleeding off the left edge: bezel, screen, then the base deck
    s.addShape('rect', { x: -1.402, y: 1.167, w: 5.874, h: 3.398, fill: { color: '111111' } });
    photo(s, { x: -0.683, y: 1.379, w: 5.04, h: 3.186 });
    addPolygon(s, [[-1.402, 4.565], [4.49, 4.565], [5.02, 4.79], [-1.402, 4.79]], '333333');
    s.addShape('roundRect', { x: -1.402, y: 4.79, w: 6.488, h: 0.102, rectRadius: 0.051, fill: { color: '6E6D6C' } });

    text(s, 'Get In Touch', { x: 5.268, y: 0.922, w: 3.839, h: 0.682, font: SEMI, size: 36, bold: true });
    CONTACTS.forEach(c => {
        const ink = c.dark ? WHITE : INK;
        if (c.dark) gradCard(s, { x: 5.367, y: c.y, w: 3.839, h: 0.863, shadow: true });
        else whiteCard(s, { x: 5.367, y: c.y, w: 3.839, h: 0.863 });
        text(s, c.label, { x: 5.542, y: c.y + 0.131, w: 1.2, h: 0.278, font: MED, size: 12, color: ink });
        text(s, L_AMET, { x: 5.542, y: c.y + 0.389, w: 3.151, h: 0.348, lh: 1.8, color: ink });
    });
    pageNumber(s, 19);
}

function slide20(pptx) {
    const s = pptx.addSlide();
    ring(s, { x: -1.24, y: -1.029, d: 3.842 });
    blob(s, { x: 5.852, y: 4.427, d: 2.161 });
    photo(s, { x: 0.186, y: 0.185, w: 3.643, h: 3.643, shape: 'ellipse' });
    photo(s, { x: 7.653, y: 3.279, w: 2.161, h: 2.161, shape: 'ellipse' });

    text(s, 'Thank You', { x: 4.292, y: 0.566, w: 5.708, h: 3.559, font: SEMI, size: 104, bold: true });
    labeled(s, { x: 0.253, y: 4.616, w: 5.607, body: L_ULLAMCO });
    pageNumber(s, 20);
}

/* ==================================================================== build */
const SLIDES = [
    slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

function build() {
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'DECK_16X9', width: 10, height: 5.625 });
    pptx.layout = 'DECK_16X9';
    pptx.title = 'Leadership Development';

    SLIDES.forEach(fn => fn(pptx));

    const out = path.join(__dirname, path.basename(__filename, '.js') + '.pptx');
    return pptx.writeFile({ fileName: out }).then(() => console.log('wrote', out));
}

build();
