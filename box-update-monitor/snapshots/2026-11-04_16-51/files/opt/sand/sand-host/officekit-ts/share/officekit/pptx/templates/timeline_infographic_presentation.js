/**
 * "Timeskip" — Timeline Infographic Presentation Template
 * A standalone pptxgenjs recreation of the 30-slide reference deck (10 x 5.625 in).
 *
 *   node 078facdd-16a0-4e3f-ad7c-876d5754e34a_grok_final.js
 *
 * The original deck embeds one cover photograph and ~70 small glyph PNGs.
 * Those are redrawn here as plain coloured placeholder shapes; everything else
 * (chevrons, rings, ribbons, gradients, type) is built from native shapes.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const OUTPUT = '078facdd-16a0-4e3f-ad7c-876d5754e34a_grok_final.pptx';

/* =========================================================== 1. PALETTE == */
const NAVY = '205072', TEAL = '329D9C', MINT = '56C596', GREEN = '60DE7E',
      PALE = 'AAEBAF', LEAF = '7BE495',
      BLUE = '539CCF', AQUA = '75D2D1', SEA = '98DCBF', LIME = '9EEBB0', MOSS = 'AEEEBE',
      WHITE = 'FFFFFF', BLACK = '000000',
      GRAY = '7F7F7F', DGRAY = '595959', LGRAY = 'F2F2F2', XGRAY = 'F3F3F3',
      SHADE = 'D8D8D8', PHOTO = '787878';

/** the deck's signature four-stop navy -> green ramp */
const RAMP  = [[0, NAVY], [0.3, TEAL], [0.7, MINT], [1, GREEN]];
/**
 * Every gradient in the reference is stretched to twice its shape's extent, so
 * only the middle half of RAMP is ever visible (teal through mint). RAMP_MID
 * reproduces that, and RAMP_DOT is the single colour a tiny shape collapses to.
 */
const RAMP_MID = (t) => rampAt(RAMP, 0.5 + (t - 0.5) / 2);
const RAMP_DOT = () => rampAt(RAMP, 0.5);
/** discrete five- and six-step versions of the same ramp */
const RAMP5 = [NAVY, TEAL, MINT, GREEN, PALE];
const RAMP6 = [NAVY, TEAL, MINT, GREEN, PALE, LEAF];

/* ============================================== 2. REPEATED BODY COPY ==== */
const L_CARD  = 'Lorem ipsum dolor sit amet, consectetu adipiscing elit, sed does eiusmod.';
const L_SHORT = 'Lorem ipsum dolor sit amet, consectetu adipiscing elit, sed.';
const L_TINY  = 'Lorem ipsum dolor sit amet, consect adipi.';
const L_ROAD  = 'Lorem ipsum dolor sit amet, consec adipiscing elit sed do eiusmod tem.';
const L_COL   = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut.';
const L_PILL  = 'Lorem ipsum dolor sit amet, consect adipisci elit, sed does.';
const L_RIBB  = 'Lorem ipsum dolor sit amet, consect adipiscing elit, sed does eiusmod.';
const L_SIDE  = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do.';
const L_BAR   = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed.';
const L_LEAD  = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehend.';
const L_LEAD2 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip.';
const L_COPY  = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sedo do eiusmod tempor incididunt ut labore et dolore magna aliq. Ut enim ad minim veniam, quis nostrud exercitation ullamco.';
const L_COPY2 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sedo do eiusmod tempor.';
const L_PROJ  = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eius tempor incididunt ut labore et dolore magna aliqua. Ut enim ad mini veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure.';
const L_PROJ2 = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eius tempor incididunt ut labore et dolore magna aliqua.';
const YTH     = 'Your Title Here';

/* ============================================== 3. SHAPE / TEXT HELPERS == */
const NL = { type: 'none' };                     // "no outline"
const F  = (color) => ({ color });               // solid fill

const shp = (s, type, x, y, w, h, o) => s.addShape(type, Object.assign({ x, y, w, h, line: NL }, o));
const box = (s, x, y, w, h, color) => shp(s, 'rect', x, y, w, h, { fill: F(color) });
const dot = (s, cx, cy, d, color) => shp(s, 'ellipse', cx - d / 2, cy - d / 2, d, d, { fill: F(color) });
const dash = (s, x, y, w, h, color, kind) =>
    s.addShape('line', { x, y, w, h, line: { color, width: 1.5, dashType: kind || 'dash' } });

/** soft card shadow used by the deck's white panels */
const CARD_SHADOW = { type: 'outer', blur: 10, offset: 2, angle: 90, color: 'BFBFBF', opacity: 0.5 };

/**
 * custGeom from points in 0..1 shape space.
 *   [x, y]                      line-to (or move-to for the first point)
 *   [cx1, cy1, cx2, cy2, x, y]  cubic bezier
 *   []                          close path
 */
function poly(s, x, y, w, h, pts, o) {
    const points = pts.map(p => p.length === 0 ? { close: true }
        : p.length === 6 ? { curve: { type: 'cubic', x1: p[0] * w, y1: p[1] * h, x2: p[2] * w, y2: p[3] * h }, x: p[4] * w, y: p[5] * h }
        : { x: p[0] * w, y: p[1] * h });
    s.addShape('custGeom', Object.assign({ x, y, w, h, points, line: NL }, o));
}

/** absolute-coordinate polygon: pts are [x, y] pairs in inches */
function polyAbs(s, pts, o) {
    const xs = pts.filter(p => p.length).map(p => p[0]), ys = pts.filter(p => p.length).map(p => p[1]);
    const x = Math.min.apply(null, xs), y = Math.min.apply(null, ys);
    const w = Math.max.apply(null, xs) - x, h = Math.max.apply(null, ys) - y;
    poly(s, x, y, w, h, pts.map(p => p.length ? [(p[0] - x) / w, (p[1] - y) / h] : p), o);
}

const RAD = (deg) => deg * Math.PI / 180;

/**
 * Annular band (a "ring segment"), swept from a0 to a1 degrees.
 * Angles run clockwise from 3 o'clock, matching screen coordinates.
 */
function arcBand(s, cx, cy, rx, ry, t, a0, a1, color) {
    const N = 28, pts = [];
    for (let i = 0; i <= N; i++) { const a = RAD(a0 + (a1 - a0) * i / N); pts.push([cx + rx * Math.cos(a), cy + ry * Math.sin(a)]); }
    for (let i = N; i >= 0; i--) { const a = RAD(a0 + (a1 - a0) * i / N); pts.push([cx + (rx - t) * Math.cos(a), cy + (ry - t) * Math.sin(a)]); }
    pts.push([]);
    polyAbs(s, pts, { fill: F(color) });
}
const ring = (s, cx, cy, d, t, color) => arcBand(s, cx, cy, d / 2, d / 2, t, 0, 360, color);

/** filled triangular arrow head of width `w`, pointing along `deg` */
function arrowHead(s, cx, cy, w, deg, color) {
    const a = RAD(deg), n = a + Math.PI / 2, L = w * 0.85;
    polyAbs(s, [
        [cx + L * Math.cos(a), cy + L * Math.sin(a)],
        [cx + w * Math.cos(n) / 2, cy + w * Math.sin(n) / 2],
        [cx - w * Math.cos(n) / 2, cy - w * Math.sin(n) / 2], []], { fill: F(color) });
}

/**
 * Paints a linear gradient as `n` adjacent slices.
 * `edge(t)` gives the shape's [start, end] extent (0..1) across the other axis,
 * which is what lets the same routine fill discs, rounded cards and plain rails.
 */
function gradShape(s, x, y, w, h, stops, edge, o) {
    o = o || {};
    const n = o.bands || 22, K = 5, horiz = !!o.horizontal, over = o.overlap === undefined ? 0.4 : o.overlap;
    for (let i = 0; i < n; i++) {
        const t0 = i / n, t1 = Math.min(1, (i + 1) / n + over / n), pts = [];
        for (let k = 0; k <= K; k++) { const t = t0 + (t1 - t0) * k / K; pts.push([t, edge(t)[1]]); }
        for (let k = K; k >= 0; k--) { const t = t0 + (t1 - t0) * k / K; pts.push([t, edge(t)[0]]); }
        pts.push([]);
        const u = (i + 0.5) / n;
        poly(s, x, y, w, h, pts.map(p => p.length ? (horiz ? p : [p[1], p[0]]) : p),
            { fill: F(typeof stops === 'function' ? stops(u) : rampAt(stops, u)) });
    }
}
const FLAT = () => [0, 1];
/** cross-axis profile of a rounded rectangle whose corner radius is `r` inches */
const roundEdge = (along, across, r) => (t) => {
    const k = r / along, d = t < k ? (k - t) / k : t > 1 - k ? (t - (1 - k)) / k : 0;
    const inset = (r / across) * (1 - Math.sqrt(Math.max(0, 1 - d * d)));
    return [inset, 1 - inset];
};
/**
 * Silhouette of a cover disc: a flat back plus one cubic bulge that reaches the
 * full width at mid-height, mirrored below. `p` = [x0, c1x, c1y, c2x, c2y],
 * i.e. the cubic (x0,0) -> (1,0.5) taken straight from the reference artwork.
 */
function discEdge(p, flip) {
    const bez = (u, a, b, c, d) => { const v = 1 - u; return v * v * v * a + 3 * v * v * u * b + 3 * v * u * u * c + u * u * u * d; };
    return (t) => {
        const target = t > 0.5 ? 1 - t : t;
        let lo = 0, hi = 1;                                     // invert the cubic in y
        for (let i = 0; i < 20; i++) { const m = (lo + hi) / 2; if (bez(m, 0, p[2], p[4], 0.5) < target) lo = m; else hi = m; }
        const e = bez((lo + hi) / 2, p[0], p[1], p[3], 1);
        return flip ? [1 - e, 1] : [0, e];
    };
}
const DISC_WHITE = [0.6460, 0.8596, 0.1127, 1, 0.2942];
const DISC_RAMP  = [0.5630, 0.8267, 0.1127, 1, 0.2942];
const DISC_PHOTO = [0.2264, 0.6536, 0.0000, 1, 0.2239];

/** interpolate a [position, hex] stop list at t = 0..1 */
function rampAt(stops, t) {
    t = Math.max(0, Math.min(1, t));
    let a = stops[0], b = stops[stops.length - 1];
    for (let i = 0; i < stops.length - 1; i++) {
        if (t >= stops[i][0] && t <= stops[i + 1][0]) { a = stops[i]; b = stops[i + 1]; break; }
    }
    const k = b[0] === a[0] ? 0 : (t - a[0]) / (b[0] - a[0]);
    let out = '';
    for (let i = 0; i < 3; i++) {
        const ca = parseInt(a[1].substr(i * 2, 2), 16), cb = parseInt(b[1].substr(i * 2, 2), 16);
        out += Math.round(ca + (cb - ca) * k).toString(16).padStart(2, '0').toUpperCase();
    }
    return out;
}

/* ---------------------------------------------------------------- text -- */
const TXT = { fontFace: 'Lato', margin: [5.4, 5.4, 2.7, 2.7], valign: 'top', wrap: true };
const txt = (s, str, o) => s.addText(str, Object.assign({}, TXT, o));

const body  = (s, str, x, y, w, h, o) => txt(s, str, Object.assign({ x, y, w, h, fontSize: 8, color: GRAY, lineSpacingMultiple: 1.5, align: 'center' }, o));
const label = (s, str, x, y, w, o)    => txt(s, str, Object.assign({ x, y, w, h: 0.252, fontSize: 11, bold: true, align: 'center' }, o));
const yearT = (s, str, x, y, w, o)    => txt(s, str, Object.assign({ x, y, w, h: 0.48, fontSize: 24, bold: true, align: 'center' }, o));
const yearC = (s, str, x, y, w, h, color, size) =>
    txt(s, str, { x, y, w, h, fontSize: size, bold: true, color, align: 'center', valign: 'middle', margin: 0 });

/** the deck's ubiquitous bold caption over a grey paragraph */
function cardText(s, x, y, w, title, color, o) {
    o = o || {};
    const al = o.align || 'center';
    txt(s, title, { x, y, w: o.tw || w, h: o.th || 0.252, fontSize: o.ts || 11, bold: true, color, align: al, fontFace: 'Lato', margin: TXT.margin, valign: 'top' });
    body(s, o.text || L_CARD, x, y + (o.gap || 0.237), w, o.bh || 0.673, { align: al, color: o.textColor || GRAY });
}

/** centred slide heading with an optional intro paragraph */
function heading(s, str, y, sub, subY, subH) {
    txt(s, str, { x: 2.47, y, w: 5.061, h: 0.48, fontSize: 24, bold: true, color: BLACK, fontFace: 'Roboto', align: 'center' });
    if (sub) body(s, sub, 2.222, subY, 5.555, subH || 0.673);
}

/** left/right "story" column: big title, sub-heading and two paragraphs */
function copyBlock(s, o) {
    txt(s, o.title, { x: o.x, y: o.y, w: o.w, h: 0.884, fontSize: 24, bold: true, color: BLACK, fontFace: 'Roboto', align: 'left' });
    label(s, YTH, o.x, o.yth, 2.44, { align: 'left', color: BLACK });
    body(s, o.p1 || L_COPY, o.x, o.p1y, o.w, 0.881, { align: 'left', color: o.color || DGRAY });
    body(s, o.p2 || L_COPY2, o.x, o.p2y, o.w, 0.465, { align: 'left', color: o.color || DGRAY });
}
/** the two vertical rhythms copyBlock takes across the deck */
const copyA = (s, x, title) => copyBlock(s, { x, y: 1.389, w: 2.731, yth: 3.556, p1y: 2.472, p2y: 3.771, title });
const copyB = (s, x, title) => copyBlock(s, { x, y: 1.389, w: 2.731, yth: 2.484, p1y: 2.683, p2y: 3.771, title });

/** icon placeholder standing in for the original glyph PNGs */
function icon(s, cx, cy, size, color) {
    const d = size * 0.66;
    shp(s, 'roundRect', cx - d / 2, cy - d / 2, d, d, { fill: F(color), rectRadius: d * 0.24 });
}

/** chevron ribbon segment: notch on the left, point on the right (or mirrored) */
function chevron(s, x, y, w, h, color, back) {
    const n = 0.059;
    poly(s, x, y, w, h, back
        ? [[1, 0], [n, 0], [0, 0.5], [n, 1], [1, 1], [1 - n, 0.5], []]
        : [[1 - n, 0], [0, 0], [n, 0.5], [0, 1], [1 - n, 1], [1, 0.5], []], { fill: F(color) });
}

/* ================================================== 4. SLIDE BUILDERS ==== */

/* --- 1 & 30: the cover pair, mirrored ----------------------------------- */
function coverSlide(s, o) {
    const m = o.mirror, X = (x, w) => m ? 10 - x - w : x;

    gradShape(s, X(0, 4.177), 0, 4.177, 5.625, [[0, WHITE], [1, WHITE]], discEdge(DISC_WHITE, m), { bands: 48 });
    gradShape(s, X(0, 3.867), 0, 3.867, 5.625, RAMP_MID, discEdge(DISC_RAMP, m), { bands: 48 });
    gradShape(s, X(0, 3.635), 0, 3.635, 5.625, [[0, PHOTO], [1, PHOTO]], discEdge(DISC_PHOTO, m), { bands: 48 });
    txt(s, '[image]', { x: X(0.6, 2.2), y: 2.65, w: 2.2, h: 0.3, fontSize: 9, color: WHITE, align: 'center' });

    txt(s, o.title, { x: X(5.251, 3.675), y: 1.862, w: 3.675, h: 1.085, fontSize: o.size, bold: true, color: BLACK, fontFace: 'Roboto', align: 'left' });
    txt(s, 'Timeline Infographic Presentation Template', { x: X(5.286, 3.542), y: 2.967, w: 3.542, h: 0.278, fontSize: 12, color: TEAL, fontFace: 'Roboto', align: 'left' });
    body(s, 'PLACEHOLDER',
        X(5.329, 3.3), 3.293, 3.3, 0.47, { align: 'left' });

    txt(s, 'Company Presentation', { x: X(7.92, 1.613), y: 0.289, w: 1.613, h: 0.227, fontSize: 9, color: NAVY, align: m ? 'left' : 'right' });
    box(s, X(3.811, 4.213), 0.4, 4.213, 0.012, 'D2DCE3');
    box(s, X(3.811, 4.353), 5.19, 4.353, 0.012, 'D2DCE3');

    ['in', '\u2709', 'f'].forEach((glyph, i) => {
        const x = X(8.499 + i * 0.3915, 0.252);
        shp(s, 'roundRect', x, 5.094, 0.252, 0.242, { fill: F(MINT), rectRadius: 0.024 });
        yearC(s, glyph, x, 5.094, 0.252, 0.242, WHITE, 9);
    });
}
const slide01 = (s) => coverSlide(s, { title: 'Timeskip',  size: 60, mirror: false });
const slide30 = (s) => coverSlide(s, { title: 'Thank you', size: 54, mirror: true });

/* --- 2: vertical arrow spine with alternating cards --------------------- */
function slide02(s) {
    txt(s, 'Welcome to Timeskip Infographic', { x: 0.705, y: 1.36, w: 2.545, h: 1.287, fontSize: 24, bold: true, color: BLACK, fontFace: 'Roboto', align: 'left' });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipisc elit, sed does eiusmod tempor incididunt ut labo et dolore magna aliqua. Ut enim ad minim venia, quis nostrud exercitation.', 0.705, 2.865, 2.545, 0.881, { align: 'left' });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipisc elit, sed does eiusmod tempor.', 0.705, 3.8, 2.545, 0.465, { align: 'left' });

    const rows = [
        { y: 4.455, c: NAVY,  year: '2018', name: 'Timeline One',   right: false },
        { y: 3.223, c: TEAL,  year: '2020', name: 'Timeline Two',   right: true  },
        { y: 1.992, c: MINT,  year: '2022', name: 'Timeline Three', right: false },
        { y: 0.772, c: GREEN, year: '2024', name: 'Timeline Four',  right: true  },
    ];
    rows.forEach(r => shp(s, 'homePlate', 6.075, r.y, 1.428, 0.398, { rotate: -90, fill: F(r.c) }));
    rows.forEach(r => {
        const cy = r.y + 0.199;
        dash(s, r.right ? 6.554 : 6.062, cy, 0.769, 0, r.c);
        shp(s, 'ellipse', 6.491, cy - 0.302, 0.604, 0.604, { fill: F(r.c), line: { color: WHITE, width: 1.5 } });
        icon(s, 6.793, cy, 0.375, WHITE);

        const cx = r.right ? 7.323 : 3.862, tx = r.right ? 7.465 : 4.005;
        shp(s, 'roundRect', cx, cy - 0.485, 2.402, 0.97, { fill: F(r.c), rectRadius: 0.162 });
        label(s, r.name, tx, cy - 0.372, 2.116, { color: WHITE, align: 'left' });
        body(s, L_CARD, tx, cy - 0.135, 2.203, 0.465, { align: 'left', color: WHITE });
        yearT(s, r.year, r.right ? 5.226 : 7.323, cy - 0.24, 1.038, { color: r.c, fontFace: 'Roboto', align: r.right ? 'right' : 'left' });
    });
}

/* --- 3: horizontal chevron band, captions above and below --------------- */
function slide03(s) {
    heading(s, 'Our Timeline Infographic', 0.463,
        'Lorem ipsum dolor sit amet, consectetur adipisc elit, sed does eiusmod tempor incididunt ut labo et dolore magna aliqua. Ut enim ad minim venia, quis nostrud exercitation.', 1.062, 0.465);

    const items = [
        { x: 0.000, year: '2020', name: 'Timeline One',   yearUp: true  },
        { x: 1.947, year: '2021', name: 'Timeline Two',   yearUp: false },
        { x: 3.897, year: '2022', name: 'Timeline Three', yearUp: true  },
        { x: 5.844, year: '2023', name: 'Timeline Four',  yearUp: false },
        { x: 7.794, year: '2024', name: 'Timeline Five',  yearUp: true  },
    ];
    for (let i = 4; i >= 0; i--) shp(s, 'homePlate', items[i].x, 3.401, 2.206, 0.398, { fill: F(RAMP5[i]) });
    items.forEach((it, i) => {
        const c = RAMP5[i], cx = it.x + 1.103;
        shp(s, 'ellipse', cx - 0.302, 3.298, 0.604, 0.604, { fill: F(c), line: { color: WHITE, width: 1.5 } });
        icon(s, cx, 3.6, 0.375, WHITE);
        shp(s, 'triangle', cx - 0.087, 3.034, 0.174, 0.15, { fill: F(c) });
        shp(s, 'triangle', cx - 0.087, 4.026, 0.174, 0.15, { fill: F(c), rotate: 180 });
        yearT(s, it.year, cx - 0.519, it.yearUp ? 2.479 : 4.252, 1.038, { color: c, fontFace: 'Roboto' });
        cardText(s, cx - 0.831, it.yearUp ? 4.252 : 2.048, 1.662, it.name, c);
    });
}

/* --- 4: one wide chevron ribbon on a white plate ------------------------ */
function slide04(s) {
    heading(s, 'Project Timeline Overview', 0.559, L_LEAD, 1.158);
    shp(s, 'chevron', 0, 2.245, 9.984, 1.417, { fill: F(WHITE), shadow: CARD_SHADOW });

    const names = ['Timeline One', 'Timeline Two', 'Timeline Three', 'Timeline Four', 'Timeline Five'];
    const years = ['2020', '2021', '2022', '2023', '2024'];
    for (let i = 4; i >= 0; i--) {
        const x = 0.177 + i * 1.8098;
        shp(s, i === 0 ? 'chevron' : 'homePlate', x, 2.32, i === 0 ? 2.443 : 2.479, 1.267, { fill: F(RAMP5[i]) });
    }
    for (let i = 0; i < 5; i++) {
        const x = 0.177 + i * 1.8098;
        icon(s, x + 1.221, 2.773, 0.6, WHITE);
        label(s, names[i], x + 0.603, 3.206, 1.201, { color: WHITE });
        yearT(s, years[i], x + 0.684, 3.871, 1.038, { color: RAMP5[i], fontFace: 'Roboto' });
        body(s, L_CARD, x + 0.372, 4.393, 1.662, 0.673);
    }
}

/* --- 5: stacked chevron ribbons feeding a dashed spine ------------------ */
function slide05(s) {
    copyBlock(s, { x: 5.792, y: 1.435, w: 3.179, yth: 2.513, p1y: 2.775, p2y: 3.725, title: 'Company Timeline Infographic',
        p1: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sedo do eiusmod tempor incididunt ut labore et dolore magna aliq. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex.',
        p2: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sedo do eiusmod tempor incididunt ut labore et dolore.' });

    dash(s, 4.455, 0, 0, 5.625, PALE, 'lgDash');
    const rows = [
        { y: 0.723, year: '2024', c: GREEN, tail: LIME },
        { y: 1.825, year: '2022', c: MINT,  tail: SEA  },
        { y: 2.928, year: '2021', c: TEAL,  tail: AQUA },
        { y: 4.030, year: '2019', c: NAVY,  tail: BLUE },
    ];
    rows.forEach(r => {
        poly(s, 1.419, r.y, 2.486, 0.872, [[0.89, 1], [0, 1], [0, 0], [0.89, 0], [1, 0.51], []], { fill: F(r.c) });
        poly(s, 3.601, r.y, 0.426, 0.872, [[0.571, 0], [0, 0], [0.429, 0.51], [0, 1], [0.571, 1], [1, 0.51], []], { fill: F(r.tail) });
        poly(s, 0.593, r.y, 1.040, 0.872, [[0.824, 1], [0, 1], [0, 0], [0.824, 0], [1, 0.51], []], { fill: F(LGRAY) });
        txt(s, r.year, { x: 0.763, y: r.y + 0.272, w: 0.699, h: 0.328, fontSize: 15, bold: true, color: r.c, align: 'center' });
        label(s, 'Timeline One', 1.72, r.y + 0.106, 1.895, { color: WHITE, align: 'left' });
        body(s, L_RIBB, 1.72, r.y + 0.301, 2.042, 0.465, { align: 'left', color: WHITE });
        dash(s, 4.027, r.y + 0.442, 0.359, 0, r.c);
        dot(s, 4.455, r.y + 0.442, 0.138, RAMP_DOT());
    });
}

/* --- 6: downward banners joined by arrows ------------------------------- */
function slide06(s) {
    heading(s, 'Business Events Journey', 0.335,
        'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolo.', 0.851, 0.257);

    // banner = flat-topped block whose foot dips to a point
    const BANNER = [[0, 0], [1, 0], [1, 0.89], [0.49, 1], [0, 0.89], []];
    const TAIL   = [[0, 0], [0.49, 0.429], [1, 0], [1, 0.571], [0.49, 1], [0, 0.571], []];
    [{ x: 1.151, top: 1.387, year: '2019', c: NAVY,  tail: BLUE, name: 'Timeline One',   low: false },
     { x: 3.396, top: 2.001, year: '2020', c: TEAL,  tail: AQUA, name: 'Timeline Two',   low: true  },
     { x: 5.641, top: 1.387, year: '2022', c: MINT,  tail: SEA,  name: 'Timeline Three', low: false },
     { x: 7.886, top: 2.001, year: '2024', c: GREEN, tail: LIME, name: 'Timeline Four',  low: true  }].forEach(col => {
        poly(s, col.x, col.top + 1.653, 0.964, 0.426, TAIL, { fill: F(col.tail) });
        poly(s, col.x, col.top + 0.827, 0.964, 1.040, BANNER, { fill: F(col.c) });
        poly(s, col.x, col.top,         0.964, 1.040, BANNER, { fill: F(LGRAY) });
        txt(s, col.year, { x: col.x + 0.111, y: col.top + 0.356, w: 0.742, h: 0.379, fontSize: 18, bold: true, color: col.c, align: 'center' });
        icon(s, col.x + 0.482, col.top + 1.347, 0.465, WHITE);
        cardText(s, col.x - 0.349, col.low ? 4.379 : 3.766, 1.662, col.name, col.c);
    });
    [2.354, 4.599, 6.844].forEach(x => shp(s, 'rightArrow', x, 2.535, 0.802, 0.398, { fill: F(NAVY) }));
}

/* --- 7: vertical snake of stadium pills --------------------------------- */
function slide07(s) {
    copyA(s, 0.837, 'Timeline of Progress');

    const rows = [
        { y: 4.2275, c: NAVY,  year: '2019', name: 'Timeline One',   left: true  },
        { y: 2.9425, c: TEAL,  year: '2020', name: 'Timeline Two',   left: false },
        { y: 1.6545, c: MINT,  year: '2022', name: 'Timeline Three', left: true  },
        { y: 0.3705, c: GREEN, year: '2024', name: 'Timeline Four',  left: false },
    ];
    // pale elbows that run off the top and bottom edges
    arcBand(s, 6.316, 0.155, 1.118, 0.663, 0.162, 90, 180, PALE);
    arcBand(s, 6.934, 5.465, 1.118, 0.663, 0.162, 270, 360, PALE);

    rows.forEach(r => shp(s, 'roundRect', r.left ? 5.5465 : 4.7235, r.y, 3.233, 1.016, { fill: F(r.c), rectRadius: 0.508 }));
    rows.forEach(r => {
        const x = r.left ? 5.5465 : 4.7235;
        const cx = r.left ? x + 0.508 : x + 2.725, cy = r.y + 0.508;
        ring(s, cx, cy, 1.433, 0.148, r.c);
        dot(s, cx, cy, 0.875, WHITE);
        icon(s, cx, cy, 0.465, r.c);

        const al = r.left ? 'left' : 'right';
        label(s, r.name, r.left ? 6.59 : 5.657, r.y + 0.162, 1.253, { color: WHITE, align: al });
        body(s, L_ROAD, r.left ? 6.59 : 4.923, r.y + 0.389, r.left ? 1.951 : 1.986, 0.465, { align: al, color: WHITE });
        txt(s, r.year, { x: r.left ? 4.103 : 8.35, y: r.y + 0.268, w: r.left ? 1.049 : 0.97, h: 0.48, fontSize: 24, bold: true, color: r.c, align: r.left ? 'right' : 'left' });
    });
}

/* --- 8: interlocking rings behind white pills --------------------------- */
function slide08(s) {
    txt(s, 'Historical Milestones Timeline', { x: 4.656, y: 0.483, w: 4.822, h: 0.48, fontSize: 24, bold: true, color: BLACK, fontFace: 'Roboto', align: 'left' });
    label(s, YTH, 0.61, 4.43, 2.44, { align: 'left', color: BLACK });
    body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sedo do eiusmod tempor incididunt ut labore.', 0.61, 4.691, 3.179, 0.465, { align: 'left', color: DGRAY });

    // four overlapping rings trace a lazy "S" from top-left to bottom-right
    const nodes = [
        { cx: 2.9865, cy: 2.1155, c: GREEN, year: '2024', tx: 0.502, right: false },
        { cx: 4.3755, cy: 2.1155, c: MINT,  year: '2022', tx: 4.910, right: true  },
        { cx: 4.3395, cy: 3.4945, c: TEAL,  year: '2020', tx: 1.854, right: false },
        { cx: 5.7595, cy: 3.4945, c: NAVY,  year: '2018', tx: 6.294, right: true  },
    ];
    box(s, 3.168, 0, 0.206, 1.14, GREEN);                                   // tail off the top
    arcBand(s, 2.981, 1.14, 0.393, 0.393, 0.206, 0, 90, GREEN);
    box(s, 5.349, 4.485, 0.206, 1.14, NAVY);                                // tail off the bottom
    arcBand(s, 5.742, 4.092, 0.393, 0.393, 0.206, 90, 180, NAVY);
    nodes.forEach(n => ring(s, n.cx, n.cy, 1.588, 0.206, n.c));
    nodes.forEach(n => shp(s, 'roundRect', n.right ? n.cx - 0.508 : n.cx + 0.508 - 3.233, n.cy - 0.508, 3.233, 1.016,
        { fill: F(WHITE), rectRadius: 0.508, shadow: CARD_SHADOW }));
    nodes.forEach(n => {
        dot(s, n.cx, n.cy, 0.875, n.c);
        icon(s, n.cx, n.cy, 0.465, WHITE);
        const al = n.right ? 'left' : 'right';
        txt(s, n.year, { x: n.right ? n.tx : n.tx + 0.698, y: n.cy - 0.3455, w: 1.253, h: 0.303, fontSize: 14, bold: true, color: n.c, align: al });
        body(s, L_ROAD, n.tx, n.cy - 0.1185, 1.951, 0.465, { align: al });
    });
}

/* --- 9: gradient rail with rounded cards -------------------------------- */
function slide09(s) {
    heading(s, 'Chronological Milestones', 0.455, L_LEAD, 1.054);
    gradShape(s, 0, 3.018, 10, 0.037, RAMP_MID, FLAT, { horizontal: true, bands: 40 });

    const cols = [
        { x: 0.588, c: NAVY,  year: '2019', name: 'Timeline One'   },
        { x: 2.428, c: TEAL,  year: '2020', name: 'Timeline Two'   },
        { x: 4.225, c: MINT,  year: '2022', name: 'Timeline Three' },
        { x: 6.027, c: GREEN, year: '2024', name: 'Timeline Four'  },
        { x: 7.823, c: MINT,  year: '2022', name: 'Timeline Three' },
    ];
    cols.forEach(col => {
        const cx = col.x + 0.7945;
        shp(s, 'ellipse', cx - 0.195, 2.842, 0.39, 0.39, { fill: F(WHITE), shadow: CARD_SHADOW });
        shp(s, 'roundRect', col.x, 3.377, 1.589, 1.985, { fill: F(col.c), rectRadius: 0.265 });
        yearT(s, col.year, cx - 0.5245, 2.233, 1.049, { color: col.c });
        dot(s, cx, 3.037, 0.244, col.c);
        icon(s, cx, 3.841, 0.465, WHITE);
        label(s, col.name, col.x + 0.215, 4.226, 1.159, { color: WHITE });
        body(s, L_PILL, col.x + 0.171, 4.457, 1.246, 0.673, { color: WHITE });
    });
}

/* --- 10: four full-bleed colour columns --------------------------------- */
function slide10(s) {
    heading(s, 'Our Journey Through Time', 0.627, L_LEAD, 1.226);
    dash(s, 0, 2.711, 10, 0, NAVY, 'lgDash');

    [{ c: NAVY,  year: '2019', name: 'Timeline One'   },
     { c: TEAL,  year: '2020', name: 'Timeline Two'   },
     { c: MINT,  year: '2022', name: 'Timeline Three' },
     { c: GREEN, year: '2024', name: 'Timeline Four'  }].forEach((col, i) => {
        const x = i * 2.5;
        box(s, x, 2.875, 2.5, 2.82, col.c);
        dot(s, x + 1.25, 2.8125, 0.875, WHITE);
        icon(s, x + 1.25, 2.8125, 0.465, col.c);
        yearT(s, col.year, x + 0.725, 3.635, 1.049, { color: WHITE });
        label(s, col.name, x + 0.67, 4.302, 1.159, { color: WHITE });
        body(s, L_COL, x + 0.266, 4.533, 1.969, 0.673, { color: WHITE });
    });
}

/* --- 11: three colour blocks off a gradient rail ------------------------ */
function slide11(s) {
    copyA(s, 1.163, 'Our Milestone Highlights');

    const rows = [
        { y: 0.000, c: NAVY, year: '2020', name: 'Timeline One'  },
        { y: 1.875, c: TEAL, year: '2022', name: 'Timeline Two'  },
        { y: 3.750, c: MINT, year: '2024', name: 'Timeline Thre' },
    ];
    rows.forEach(r => {
        box(s, 6.508, r.y, 3.492, 1.875, r.c);
        dot(s, 7.177, r.y + 0.938, 0.875, WHITE);
        icon(s, 7.177, r.y + 0.938, 0.465, r.c);
        txt(s, r.year, { x: 7.845, y: r.y + 0.328, w: 1.049, h: 0.48, fontSize: 24, bold: true, color: WHITE, align: 'left' });
        label(s, r.name, 7.845, r.y + 0.851, 1.159, { color: WHITE, align: 'left' });
        body(s, L_SIDE, 7.845, r.y + 1.082, 1.859, 0.465, { align: 'left', color: WHITE });
        dash(s, 5.687, r.y + 0.938, 0.82, 0, r.c, 'lgDash');
    });
    gradShape(s, 5.474, 0, 0.037, 5.625, RAMP_MID, FLAT, { bands: 40 });
    rows.forEach(r => {
        shp(s, 'ellipse', 5.297, r.y + 0.742, 0.39, 0.39, { fill: F(WHITE), shadow: CARD_SHADOW });
        dot(s, 5.492, r.y + 0.937, 0.244, r.c);
    });
}

/* --- 12: "eye" nodes with bracketed captions ---------------------------- */
function slide12(s) {
    copyA(s, 6.384, 'Evolution of Our Best Events');

    /** open rounded bracket wrapping a caption; the gap faces the node */
    const bracket = (x, y, w, h, color, hingeLeft) => {
        const r = 0.16, pts = hingeLeft
            ? [[0.42 * w, h], [r, h], [0, h - r, 0, h - r, 0, h - 2 * r], [0, 2 * r], [0, r, 0, r, r, 0], [w, 0]]
            : [[0.58 * w, h], [w - r, h], [w, h - r, w, h - r, w, h - 2 * r], [w, 2 * r], [w, r, w, r, w - r, 0], [0, 0]];
        s.addShape('custGeom', { x, y, w, h, line: { color, width: 0.75 }, fill: NL,
            points: pts.map(p => p.length === 6
                ? { curve: { type: 'cubic', x1: p[0], y1: p[1], x2: p[2], y2: p[3] }, x: p[4], y: p[5] }
                : { x: p[0], y: p[1] }) });
    };

    // odd rows keep the node left of centre, even rows right of it
    [{ cy: 0.937, c: GREEN, nodeRight: false },
     { cy: 2.181, c: MINT,  nodeRight: true  },
     { cy: 3.439, c: TEAL,  nodeRight: false },
     { cy: 4.686, c: NAVY,  nodeRight: true  }].forEach(r => {
        const cx = r.nodeRight ? 3.769 : 2.303;
        // leader: a slim wedge either side of the node, tipped with a dot
        [-1, 1].forEach(d => {
            polyAbs(s, [[cx + d * 0.788, r.cy], [cx + d * 0.42, r.cy - 0.115], [cx + d * 0.42, r.cy + 0.115], []], { fill: F(r.c) });
            dot(s, cx + d * 0.792, r.cy, 0.069, r.c);
        });
        ring(s, cx, r.cy, 0.9, 0.09, r.c);
        ring(s, cx, r.cy, 0.72, 0.09, r.c);
        icon(s, cx, r.cy, 0.375, r.c);
        arcBand(s, cx, r.cy, 0.627, 0.627, 0.034, 105, 165, r.c);          // small swoosh accent

        // caption + bracket go opposite the node; the year sits on the far side
        const capRight = !r.nodeRight;
        bracket(capRight ? cx + 0.788 : cx - 2.042, r.cy - 0.45, 1.254, 0.901, r.c, capRight);
        const tx = capRight ? 3.247 : 0.659, al = capRight ? 'left' : 'right';
        cardText(s, tx, r.cy - 0.339, 2.14, 'Timeline Four', r.c, { align: al, tw: 1.662, bh: 0.465 });
        yearT(s, '2022', capRight ? 0.295 : 4.583, r.cy - 0.24, 1.181, { color: r.c, fontFace: 'Roboto' });
    });
}

/* --- 13: zig-zagging outline circles ------------------------------------ */
function slide13(s) {
    heading(s, ' Timeline of Achievements', 0.596, L_LEAD, 1.195);

    [[1.889, NAVY, false], [3.72, TEAL, true], [5.58, MINT, false], [7.333, GREEN, true]]
        .forEach(([x, c, flip]) => s.addShape('line', { x, y: 3.344, w: 0.942, h: 0.868, flipV: flip, line: { color: c, width: 1.5, dashType: 'lgDash' } }));

    [{ x: 0.872, y: 2.509, c: NAVY,  year: '2019', name: 'Timeline One',   up: true  },
     { x: 2.740, y: 3.859, c: TEAL,  year: '2020', name: 'Timeline Two',   up: false },
     { x: 4.527, y: 2.513, c: MINT,  year: '2022', name: 'Timeline Three', up: true  },
     { x: 6.314, y: 3.859, c: GREEN, year: '2023', name: 'Timeline Four',  up: false },
     { x: 8.102, y: 2.513, c: PALE,  year: '2024', name: 'Timeline Five',  up: true  }].forEach(n => {
        shp(s, 'ellipse', n.x, n.y, 1.11, 1.11, { fill: F(n.c) });
        dot(s, n.x + 0.555, n.y + 0.555, 0.927, WHITE);
        yearC(s, n.year, n.x + 0.09, n.y + 0.09, 0.927, 0.927, n.c, 15);
        const cx = n.x + 0.555;
        shp(s, 'triangle', cx - 0.087, n.up ? 3.714 : 3.597, 0.174, 0.15, { fill: F(n.c), rotate: n.up ? 0 : 180 });
        cardText(s, cx - 0.831, n.up ? 3.959 : 2.613, 1.662, n.name, n.c);
    });
}

/* --- 14: circles linked by swooping half-arrows ------------------------- */
function slide14(s) {
    heading(s, 'Step by Step Timeline Analysis', 0.413, L_LEAD2, 0.948, 0.465);

    const nodes = [
        { cx: 2.3425, c: NAVY,  halo: '8ABEE0', year: '2019', name: 'Timeline Two',   up: true  },
        { cx: 4.0915, c: TEAL,  halo: 'D0F0F0', year: '2020', name: 'Timeline Two',   up: false },
        { cx: 5.8395, c: MINT,  halo: 'B9E8D4', year: '2022', name: 'Timeline Three', up: true  },
        { cx: 7.5875, c: GREEN, halo: 'BEF2CA', year: '2024', name: 'Timeline Four',  up: false },
    ];
    const CY = 3.518;
    nodes.forEach((n, i) => {
        // arcs alternate below / above the row of circles
        const below = i % 2 === 0, a0 = below ? 0 : 180, a1 = below ? 180 : 360;
        arcBand(s, n.cx + 0.037, CY, 0.969, 0.939, 0.15, a0 + 12, a1 - 4, n.c);
        arrowHead(s, n.cx + 0.037 + 0.894 * (below ? 1 : -1), CY - 0.06 * (below ? 1 : -1), 0.36, below ? -70 : 110, n.c);
    });
    nodes.forEach(n => {
        dot(s, n.cx, CY, 1.083, n.halo);
        shp(s, 'ellipse', n.cx - 0.52, CY - 0.504, 1.04, 1.008, { fill: F(n.c) });
        yearC(s, n.year, n.cx - 0.52, CY - 0.504, 1.04, 1.008, WHITE, 21);
        shp(s, 'triangle', n.cx - 0.087, n.up ? 2.599 : 4.296, 0.174, 0.15, { fill: F(n.c), rotate: n.up ? 180 : 0 });
        cardText(s, n.cx - 1.0775, n.up ? 1.786 : 4.541, 2.155, n.name, n.c, { bh: 0.465 });
    });
}

/* --- 15: white cards with chevron year tabs ----------------------------- */
function slide15(s) {
    heading(s, 'Our Company History', 0.511, L_LEAD, 1.227);
    dash(s, 0, 2.673, 10, 0, NAVY, 'lgDash');

    const cols = [
        { x: 0.702, c: NAVY,  year: '2019' },
        { x: 2.937, c: TEAL,  year: '2020' },
        { x: 5.177, c: MINT,  year: '2022' },
        { x: 7.406, c: GREEN, year: '2024' },
    ];
    cols.forEach(col => shp(s, 'roundRect', col.x, 2.223, 1.891, 3.102, { fill: F(WHITE), rectRadius: 0.315, shadow: CARD_SHADOW }));
    cols.forEach(col => {
        shp(s, 'homePlate', col.x + 0.261, 2.478, 1.37, 0.391, { fill: F(col.c) });
        yearC(s, col.year, col.x + 0.261, 2.478, 1.24, 0.391, WHITE, 14);
        cardText(s, col.x + 0.115, 3.07, 1.662, 'Timeline One', col.c);
        shp(s, 'ellipse', col.x + 0.501, 4.224, 0.891, 0.891, { fill: F(WHITE), shadow: CARD_SHADOW });
        icon(s, col.x + 0.9465, 4.6695, 0.465, col.c);
    });
}

/* --- 16: alternating bars off a vertical grey rail ---------------------- */
function slide16(s) {
    copyB(s, 0.841, 'Chronological Breakdown');
    shp(s, 'roundRect', 6.862, 0.184, 0.185, 5.258, { fill: F(LGRAY), rectRadius: 0.0925 });

    [{ y: 0.537, c: GREEN, pip: LIME, year: '2024', name: 'Timeline Four',  left: true  },
     { y: 1.787, c: MINT,  pip: SEA,  year: '2022', name: 'Timeline Three', left: false },
     { y: 3.037, c: TEAL,  pip: AQUA, year: '2020', name: 'Timeline Two',   left: true  },
     { y: 4.300, c: NAVY,  pip: BLUE, year: '2019', name: 'Timeline One',   left: false }].forEach(r => {
        const bx = r.left ? 5.129 : 7.047, ex = r.left ? 4.416 : 8.446;
        box(s, bx, r.y, 1.732, 0.834, r.c);
        shp(s, 'ellipse', ex, r.y - 0.105, 1.045, 1.045, { fill: F(r.c) });
        dot(s, ex + 0.523, r.y + 0.417, 0.834, WHITE);
        icon(s, ex + 0.523, r.y + 0.417, 0.465, r.c);
        txt(s, r.year, { x: r.left ? 5.461 : 7.081, y: r.y + 0.152, w: 1.364, h: 0.53, fontSize: 27, bold: true, color: WHITE, align: 'center' });
        shp(s, 'roundRect', r.left ? 7.189 : 5.452, r.y, 1.266, 0.303, { fill: F(r.c), rectRadius: 0.1515 });
        yearC(s, r.name, r.left ? 7.189 : 5.452, r.y, 1.266, 0.303, WHITE, 11);
        body(s, L_CARD, r.left ? 7.189 : 4.597, r.y + 0.371, 2.12, 0.465, { align: r.left ? 'left' : 'right' });
        dot(s, 6.955, r.y + 0.418, 0.302, WHITE);
        dot(s, 6.955, r.y + 0.418, 0.19, r.pip);
    });
}

/* --- 17: the same bar motif turned horizontal --------------------------- */
function slide17(s) {
    shp(s, 'roundRect', 0.184, 2.752, 9.744, 0.185, { fill: F(LGRAY), rectRadius: 0.0925 });

    [{ x: 0.539, c: NAVY,  pip: BLUE, year: '2020', name: 'Timeline One',   down: true  },
     { x: 2.175, c: TEAL,  pip: AQUA, year: '2021', name: 'Timeline Two',   down: false },
     { x: 3.778, c: MINT,  pip: SEA,  year: '2022', name: 'Timeline Three', down: true  },
     { x: 5.414, c: GREEN, pip: LIME, year: '2023', name: 'Timeline Four',  down: false },
     { x: 7.017, c: LEAF,  pip: MOSS, year: '2024', name: 'Timeline Five',  down: true  }].forEach(col => {
        const y = col.down ? 4.167 : 1.498;
        box(s, col.x, y, 1.731, 0.834, col.c);
        shp(s, 'ellipse', col.x + 1.399, y - 0.105, 1.045, 1.045, { fill: F(col.c) });
        dot(s, col.x + 1.922, y + 0.417, 0.834, WHITE);
        icon(s, col.x + 1.922, y + 0.417, 0.465, col.c);
        txt(s, col.year, { x: col.x + 0.034, y: y + 0.152, w: 1.364, h: 0.53, fontSize: 27, bold: true, color: WHITE, align: 'center' });
        shp(s, 'roundRect', col.x + 0.589, col.down ? 3.187 : 0.518, 1.266, 0.303, { fill: F(col.c), rectRadius: 0.1515 });
        yearC(s, col.name, col.x + 0.589, col.down ? 3.187 : 0.518, 1.266, 0.303, WHITE, 11);
        body(s, L_CARD, col.x + 0.162, col.down ? 3.558 : 0.888, 2.12, 0.465, { align: 'right' });
        dot(s, col.x + 1.222, 2.844, 0.302, WHITE);
        dot(s, col.x + 1.222, 2.844, 0.19, col.pip);
    });
}

/* --- 18: teardrop markers straddling a grey rail ------------------------ */
function slide18(s) {
    heading(s, 'A Walk Through Time', 0.392);
    shp(s, 'roundRect', 0.87, 3.135, 8.092, 0.167, { fill: F(LGRAY), rectRadius: 0.0835 });

    [{ x: 0.392, c: NAVY,  pip: BLUE, year: '2020', name: 'Timeline One',   down: true  },
     { x: 2.344, c: TEAL,  pip: AQUA, year: '2021', name: 'Timeline Two',   down: false },
     { x: 4.290, c: MINT,  pip: SEA,  year: '2022', name: 'Timeline Three', down: true  },
     { x: 6.227, c: GREEN, pip: LIME, year: '2023', name: 'Timeline Four',  down: false },
     { x: 8.196, c: LEAF,  pip: MOSS, year: '2024', name: 'Timeline Five',  down: true  }].forEach(col => {
        const y = col.down ? 3.851 : 1.14, cx = col.x + 0.7105;
        shp(s, 'triangle', cx - 0.175, col.down ? 3.594 : 2.515, 0.35, 0.302, { fill: F(col.c), rotate: col.down ? 0 : 180 });
        shp(s, 'ellipse', col.x, y, 1.421, 1.421, { fill: F(col.c) });
        dot(s, cx, y + 0.71, 1.202, WHITE);
        icon(s, cx, y + 0.42, 0.465, col.c);
        label(s, col.name, col.x + 0.218, y + 0.768, 0.983, { color: col.c });
        txt(s, col.year, { x: cx - 0.831, y: col.down ? 1.775 : 3.539, w: 1.662, h: 0.379, fontSize: 18, bold: true, color: col.c, align: 'center' });
        body(s, L_CARD, cx - 0.831, col.down ? 2.199 : 3.963, 1.662, 0.673);
        dot(s, cx, 3.219, 0.302, WHITE);
        dot(s, cx, 3.219, 0.19, col.pip);
    });
}

/* --- 19: 2x2 card grid inside a dashed bracket -------------------------- */
function slide19(s) {
    copyBlock(s, { x: 5.693, y: 1.436, w: 3.542, yth: 2.531, p1y: 2.73, p2y: 3.724,
        title: 'Timeline of Our Project Progress', color: GRAY, p1: L_PROJ, p2: L_PROJ2 });
    shp(s, 'round2SameRect', 0.84, 0.406, 2.866, 4.812, { rotate: 90, fill: NL, line: { color: NAVY, width: 2, dashType: 'lgDash' } });

    [{ x: 0.344, y: 0.223, c: NAVY,  year: '2019', name: 'Timeline One'   },
     { x: 2.402, y: 0.223, c: TEAL,  year: '2020', name: 'Timeline Two'   },
     { x: 2.402, y: 2.961, c: MINT,  year: '2022', name: 'Timeline Three' },
     { x: 0.344, y: 2.961, c: GREEN, year: '2024', name: 'Timeline Four'  }].forEach(cd => {
        shp(s, 'roundRect', cd.x, cd.y, 1.711, 2.441, { fill: F(cd.c), rectRadius: 0.285 });
        yearT(s, cd.year, cd.x + 0.331, cd.y + 0.256, 1.049, { color: WHITE });
        icon(s, cd.x + 0.8555, cd.y + 1.0875, 0.465, WHITE);
        label(s, cd.name, cd.x + 0.276, cd.y + 1.489, 1.159, { color: WHITE });
        body(s, L_TINY, cd.x + 0.232, cd.y + 1.719, 1.246, 0.465, { color: WHITE });
    });
}

/* --- 20: two rows of fat arrows inside a dashed bracket ----------------- */
function slide20(s) {
    shp(s, 'round2SameRect', 3.683, -1.918, 2.369, 9.828, { rotate: 90, fill: NL, line: { color: NAVY, width: 2, dashType: 'lgDash' } });

    [{ x: 0.572, y: 0.876, c: NAVY,  year: '2019', name: 'Timeline One',   back: false },
     { x: 3.727, y: 0.876, c: TEAL,  year: '2020', name: 'Timeline Two',   back: false },
     { x: 6.881, y: 0.876, c: MINT,  year: '2021', name: 'Timeline Three', back: false },
     { x: 5.304, y: 3.264, c: GREEN, year: '2023', name: 'Timeline Four',  back: true  },
     { x: 2.149, y: 3.264, c: LEAF,  year: '2024', name: 'Timeline Five',  back: true  }].forEach(a => {
        shp(s, a.back ? 'leftArrow' : 'rightArrow', a.x, a.y, 2.547, 1.852, { fill: F(a.c) });
        const tx = a.back ? a.x + 0.922 : a.x;
        txt(s, a.year, { x: tx, y: a.y, w: 1.625, h: 0.379, fontSize: 18, bold: true, color: a.c, align: 'center' });
        body(s, L_TINY, a.back ? a.x + 1.061 : a.x + 0.24, a.y + 0.673, 1.246, 0.465, { align: 'right', color: WHITE });
        icon(s, a.back ? a.x + 0.66 : a.x + 1.886, a.y + 0.906, 0.465, WHITE);
        txt(s, a.name, { x: tx, y: a.y + 1.472, w: 1.625, h: 0.303, fontSize: 14, bold: true, color: a.c, align: 'center' });
    });
}

/* --- 21: tall cards over a dashed baseline ------------------------------ */
function slide21(s) {
    heading(s, 'Key Dates and Events', 0.404);
    box(s, 0, 2.102, 10, 1.781, WHITE);
    s.addShape('rect', { x: 0, y: 2.102, w: 10, h: 1.781, fill: F(WHITE), line: NL, shadow: { type: 'outer', blur: 10, offset: 2, angle: 90, color: 'D9D9D9', opacity: 0.5 } });
    dash(s, 0, 5.198, 10, 0, NAVY, 'lgDash');

    [{ x: 0.312, c: NAVY,  year: '2019', name: 'Timeline One'   },
     { x: 2.707, c: TEAL,  year: '2020', name: 'Timeline Two'   },
     { x: 5.102, c: MINT,  year: '2022', name: 'Timeline Three' },
     { x: 7.496, c: GREEN, year: '2024', name: 'Timeline Four'  }].forEach(col => {
        shp(s, 'roundRect', col.x, 1.28, 2.191, 3.424, { fill: F(col.c), rectRadius: 0.365 });
        yearT(s, col.year, col.x + 0.571, 1.602, 1.049, { color: WHITE });
        label(s, col.name, col.x + 0.516, 2.268, 1.159, { color: WHITE });
        body(s, L_COL, col.x + 0.112, 2.499, 1.969, 0.673, { color: WHITE });
        dot(s, col.x + 1.0995, 3.9365, 0.875, WHITE);
        icon(s, col.x + 1.0995, 3.9365, 0.465, col.c);
        dash(s, col.x + 1.096, 4.704, 0, 0.494, NAVY, 'lgDash');
        dot(s, col.x + 1.096, 5.198, 0.203, col.c);
    });
}

/* --- 22: gradient cards hanging off a gradient rail --------------------- */
function slide22(s) {
    copyB(s, 1.044, 'Our Historical Progression');
    dash(s, 5.095, 0, 0, 5.625, PALE, 'lgDash');

    [{ y: 0.372, year: '2020', name: 'Timeline One'   },
     { y: 2.220, year: '2022', name: 'Timeline Two'   },
     { y: 4.069, year: '2024', name: 'Timeline Three' }].forEach(r => {
        gradShape(s, 6.172, r.y - 0.253, 3.517, 1.688, RAMP_MID, roundEdge(3.517, 1.688, 0.281), { horizontal: true, bands: 26 });
        dot(s, 6.9085, r.y + 0.5915, 0.875, WHITE);
        icon(s, 6.9085, r.y + 0.5915, 0.465, TEAL);
        txt(s, r.year, { x: 7.599, y: r.y, w: 1.049, h: 0.48, fontSize: 24, bold: true, color: WHITE, align: 'left' });
        label(s, r.name, 7.599, r.y + 0.501, 1.159, { color: WHITE, align: 'left' });
        body(s, L_BAR, 7.599, r.y + 0.719, 1.807, 0.465, { align: 'left', color: WHITE });
        dash(s, 5.164, r.y + 0.592, 1.008, 0, GREEN, 'lgDash');
        dot(s, 5.095, r.y + 0.592, 0.138, RAMP_DOT());
    });
}

/* --- 23: white columns capped with clipped colour blocks ---------------- */
function slide23(s) {
    heading(s, 'Timeline of Key Events', 0.492, L_LEAD2, 1.177, 0.465);

    const cols = [
        { x: 0.219, c: NAVY,  year: '2020', name: 'Timeline One'   },
        { x: 2.133, c: TEAL,  year: '2021', name: 'Timeline Two'   },
        { x: 4.047, c: MINT,  year: '2022', name: 'Timeline Three' },
        { x: 5.961, c: GREEN, year: '2023', name: 'Timeline Four'  },
        { x: 7.875, c: LEAF,  year: '2024', name: 'Timeline Five'  },
    ];
    cols.forEach(col => s.addShape('rect', { x: col.x, y: 2.648, w: 1.906, h: 2.977, fill: F(WHITE), line: NL, shadow: { type: 'outer', blur: 10, offset: 2, angle: 90, color: 'D9D9D9', opacity: 0.5 } }));
    cols.forEach(col => {
        shp(s, 'snip2SameRect', col.x + 0.433, 1.691, 1.047, 1.914, { rotate: 90, fill: F(col.c) });
        yearT(s, col.year, col.x + 0.432, 2.409, 1.049, { color: WHITE });
        cardText(s, col.x + 0.126, 3.37, 1.662, col.name, col.c);
        shp(s, 'ellipse', col.x + 0.511, 4.477, 0.891, 0.891, { fill: F(col.c) });
        icon(s, col.x + 0.9565, 4.9225, 0.465, WHITE);
    });
}

/* --- 24: rounded "app tile" badges over a dashed baseline --------------- */
function slide24(s) {
    heading(s, 'Milestone Mapping', 0.46, L_LEAD2, 1.067, 0.465);
    dash(s, 0, 4.016, 10, 0, NAVY, 'lgDash');

    [{ x: 0.582, c: NAVY,  year: '2019', name: 'Timeline One'   },
     { x: 2.980, c: TEAL,  year: '2020', name: 'Timeline Two'   },
     { x: 5.375, c: MINT,  year: '2022', name: 'Timeline Three' },
     { x: 7.773, c: GREEN, year: '2024', name: 'Timeline Four'  }].forEach(col => {
        shp(s, 'roundRect', col.x, 2.118, 1.637, 1.508, { fill: F(col.c), rectRadius: 0.251 });
        shp(s, 'roundRect', col.x + 0.233, 1.922, 1.172, 1.079, { fill: F(WHITE), rectRadius: 0.18 });
        shp(s, 'roundRect', col.x + 0.287, 1.972, 1.063, 0.979, { fill: F(col.c), rectRadius: 0.163 });
        icon(s, col.x + 0.8185, 2.4615, 0.525, WHITE);
        txt(s, col.year, { x: col.x + 0.294, y: 3.097, w: 1.049, h: 0.429, fontSize: 21, bold: true, color: WHITE, align: 'center' });
        shp(s, 'triangle', col.x + 0.731, 3.69, 0.174, 0.15, { fill: F(col.c) });
        dot(s, col.x + 0.8255, 4.0155, 0.203, col.c);
        cardText(s, col.x - 0.005, 4.385, 1.662, col.name, NAVY);
    });
}

/* --- 25: the same badges stacked down a dashed spine -------------------- */
function slide25(s) {
    copyBlock(s, { x: 5.551, y: 1.436, w: 3.542, yth: 2.531, p1y: 2.73, p2y: 3.724,
        title: 'Tracking Our Company Journey', color: GRAY, p1: L_PROJ, p2: L_PROJ2 });
    dash(s, 1.72, 0, 0, 5.625, NAVY, 'lgDash');

    [{ y: 0.299, c: NAVY, year: '2020', name: 'Timeline One'   },
     { y: 2.157, c: TEAL, year: '2022', name: 'Timeline Two'   },
     { y: 4.014, c: MINT, year: '2024', name: 'Timeline Three' }].forEach(r => {
        shp(s, 'roundRect', 0.907, r.y, 1.637, 1.508, { fill: F(r.c), rectRadius: 0.251 });
        shp(s, 'roundRect', 1.14, r.y - 0.196, 1.172, 1.079, { fill: F(WHITE), rectRadius: 0.18 });
        shp(s, 'roundRect', 1.194, r.y - 0.146, 1.063, 0.979, { fill: F(r.c), rectRadius: 0.163 });
        icon(s, 1.7255, r.y + 0.3435, 0.525, WHITE);
        txt(s, r.year, { x: 1.201, y: r.y + 0.979, w: 1.049, h: 0.429, fontSize: 21, bold: true, color: WHITE, align: 'center' });
        cardText(s, 2.838, r.y + 0.201, 1.548, r.name, r.c, { align: 'left' });
    });
}

/* --- 26: three chevron rows joined by two U-turns ----------------------- */
function slide26(s) {
    /** map pin: teardrop body, white disc, glyph */
    const pin = (x, y, color) => {
        poly(s, x, y, 0.546, 0.659, [[0.854, 0.121], [0.658, -0.04, 0.342, -0.04, 0.146, 0.121],
            [-0.049, 0.283, -0.049, 0.545, 0.146, 0.707], [0.5, 1], [0.854, 0.707],
            [1.049, 0.545, 1.049, 0.283, 0.854, 0.121], []], { fill: F(color) });
        dot(s, x + 0.273, y + 0.277, 0.388, WHITE);
        icon(s, x + 0.273, y + 0.277, 0.246, color);
    };

    const top = [
        { x: 1.021, c: NAVY, year: '2017', name: 'Timeline One',   lc: NAVY },
        { x: 3.008, c: BLUE, year: '2018', name: 'Timeline Two',   lc: BLUE },
        { x: 5.000, c: TEAL, year: '2019', name: 'Timeline Three', lc: TEAL },
    ];
    const mid = [{ x: 2.984, c: SEA, year: '2021' }, { x: 4.975, c: MINT, year: '2020' }];
    const bot = [
        { x: 3.015, c: LIME, year: '2022', name: 'Timeline Four', lc: SEA   },
        { x: 5.008, c: PALE, year: '2023', name: 'Timeline Five', lc: GREEN },
        { x: 6.989, c: LEAF, year: '2024', name: 'Timeline Six',  lc: LEAF  },
    ];
    // right-hand U-turn (aqua) and left-hand U-turn (green), each 0.427 wide
    arcBand(s, 6.981, 2.032, 2.008, 0.981, 0.427, -90, 90, AQUA);
    arcBand(s, 3.139, 3.584, 2.129, 0.998, 0.427, 90, 270, GREEN);
    arcBand(s, 6.981, 2.032, 1.795, 0.768, 0.02, -84, 84, WHITE);   // white centre line
    arcBand(s, 3.139, 3.584, 1.916, 0.785, 0.02, 96, 264, WHITE);

    top.forEach(t => chevron(s, t.x, 1.051, 2.11, 0.427, t.c, true));
    mid.forEach(m => chevron(s, m.x, 2.585, 2.12, 0.427, m.c, false));
    bot.forEach(b => chevron(s, b.x, 4.155, 2.11, 0.427, b.c, true));

    top.forEach((t, i) => {
        pin(1.725 + i * 1.988, 0.231, t.c);
        txt(s, t.year, { x: t.x + 0.53, y: 1.049, w: 1.049, h: 0.429, fontSize: 21, bold: true, color: WHITE, align: 'center' });
        cardText(s, t.x + 0.225, 1.628, 1.662, t.name, t.lc, { gap: 0.303, th: 0.303, ts: 14, bh: 0.465, text: L_SHORT });
    });
    mid.forEach(m => txt(s, m.year, { x: m.x + 0.535, y: 2.583, w: 1.049, h: 0.429, fontSize: 21, bold: true, color: WHITE, align: 'center' }));
    bot.forEach((b, i) => {
        pin(3.765 + i * 1.9745, 3.335, b.c);
        txt(s, b.year, { x: b.x + 0.532, y: 4.153, w: 1.049, h: 0.429, fontSize: 21, bold: true, color: WHITE, align: 'center' });
        cardText(s, b.x + 0.225, 4.712, 1.662, b.name, b.lc, { gap: 0.303, th: 0.303, ts: 14, bh: 0.465, text: L_SHORT });
    });
}

/* --- 27: zig-zag band with stalked captions ----------------------------- */
function slide27(s) {
    [[1.727, NAVY], [7.491, LEAF]].forEach(([x, c]) => {
        shp(s, 'ellipse', x, 2.427, 0.781, 0.781, { fill: F(c) });
        dot(s, x + 0.3905, 2.8175, 0.487, XGRAY);
    });
    [{ x: 2.118, y: 2.344, c: NAVY,  up: true  },
     { x: 3.271, y: 2.597, c: TEAL,  up: false },
     { x: 4.423, y: 2.344, c: MINT,  up: true  },
     { x: 5.576, y: 2.597, c: GREEN, up: false },
     { x: 6.729, y: 2.344, c: LEAF,  up: true  }].forEach(g => poly(s, g.x, g.y, 1.153, 0.686, g.up
        ? [[1, 1], [0.5, 0.631], [0, 1], [0, 0.369], [0.5, 0], [1, 0.369], []]
        : [[0, 0], [0.5, 0.369], [1, 0], [1, 0.631], [0.5, 1], [0, 0.631], []], { fill: F(g.c) }));

    [{ x: 2.624, c: NAVY,  year: '2020', name: 'Timeline One',   up: true  },
     { x: 3.777, c: TEAL,  year: '2021', name: 'Timeline Two',   up: false },
     { x: 4.930, c: MINT,  year: '2022', name: 'Timeline Three', up: true  },
     { x: 6.083, c: GREEN, year: '2023', name: 'Timeline Four',  up: false },
     { x: 7.236, c: LEAF,  year: '2024', name: 'Timeline Five',  up: true  }].forEach(it => {
        const y = it.up ? 1.363 : 4.122;
        shp(s, 'ellipse', it.x, y, 0.14, 0.14, { fill: NL, line: { color: it.c, width: 2.34 } });
        dot(s, it.x + 0.07, y + 0.07, 0.09, it.c);
        box(s, it.x + 0.063, it.up ? 1.503 : 3.257, 0.019, 0.865, it.c);
        icon(s, it.x + 0.07, it.up ? 1.743 : 3.631, 0.643, it.c);
        txt(s, it.year, { x: it.x - 0.454, y: it.up ? 2.137 : 4.025, w: 1.049, h: 0.429, fontSize: 21, bold: true, color: it.c, align: 'center' });
        cardText(s, it.x - 0.761, it.up ? 0.403 : 4.462, 1.662, it.name, it.c,
            { gap: 0.303, th: 0.303, ts: 14, bh: 0.465, text: L_SHORT, textColor: it.c === LEAF ? LEAF : GRAY });
    });
}

/* --- 28: two slab rows joined on the left, diamond badges --------------- */
function slide28(s) {
    const DIA = [[0.584, 0.965], [0.538, 1.012, 0.462, 1.012, 0.416, 0.965], [0.035, 0.584],
        [-0.012, 0.538, -0.012, 0.462, 0.035, 0.416], [0.416, 0.035],
        [0.462, -0.012, 0.538, -0.012, 0.584, 0.035], [0.965, 0.416],
        [1.012, 0.462, 1.012, 0.538, 0.965, 0.584], []];

    // upper band bends down on the left, lower band bends up: together a "C"
    poly(s, 1.506, 2.262, 2.333, 1.252, [[0.475, 0.206], [1, 0.206], [1, 0], [0.475, 0],
        [0.213, 0, 0, 0.397, 0, 0.885], [0, 1], [0.111, 1], [0.111, 0.885],
        [0.111, 0.511, 0.274, 0.206, 0.475, 0.206], []], { fill: F(MINT) });
    poly(s, 1.506, 3.514, 2.333, 1.252, [[0, 0.115], [0, 0.603, 0.213, 1, 0.475, 1], [1, 1], [1, 0.794],
        [0.475, 0.794], [0.274, 0.794, 0.111, 0.489, 0.111, 0.115], [0.111, 0], [0, 0], []], { fill: F(GREEN) });
    box(s, 3.840, 2.262, 2.320, 0.258, TEAL);
    box(s, 6.160, 2.262, 2.333, 0.258, NAVY);
    box(s, 3.840, 4.508, 2.320, 0.258, PALE);
    box(s, 6.160, 4.508, 2.333, 0.258, LEAF);

    [{ x: 2.306, y: 2.030, c: MINT,  year: '2021', name: 'Timeline Three', ty: 0.623 },
     { x: 4.639, y: 2.030, c: TEAL,  year: '2020', name: 'Timeline Two',   ty: 0.623 },
     { x: 6.973, y: 2.030, c: NAVY,  year: '2019', name: 'Timeline One',   ty: 0.632 },
     { x: 2.306, y: 4.280, c: GREEN, year: '2022', name: 'Timeline Four',  ty: 2.896, lc: MINT  },
     { x: 4.639, y: 4.280, c: PALE,  year: '2023', name: 'Timeline Five',  ty: 2.896, lc: GREEN },
     { x: 6.973, y: 4.280, c: LEAF,  year: '2024', name: 'Timeline Six',   ty: 2.896, lc: LEAF  }].forEach(b => {
        poly(s, b.x, b.y, 0.722, 0.722, DIA, { fill: F(b.c) });
        poly(s, b.x + 0.075, b.y + 0.075, 0.571, 0.571, DIA, { fill: F(LGRAY) });
        icon(s, b.x + 0.361, b.y + 0.361, 0.246, b.c);
        const lc = b.lc || b.c;
        txt(s, b.year, { x: b.x - 0.152, y: b.ty + 0.807, w: 1.049, h: 0.429, fontSize: 21, bold: true, color: lc, align: 'center' });
        txt(s, b.name, { x: b.x - 0.348, y: b.ty, w: 1.662, h: 0.303, fontSize: 14, bold: true, color: lc, align: 'center' });
        body(s, L_SHORT, b.x - 0.348, b.ty + 0.303, 1.662, 0.465);
    });
}

/* --- 29: chain of outline circles with icon bubbles --------------------- */
function slide29(s) {
    // grey "links" arcing between neighbouring circles
    [[2.407, 2.815, 'b'], [4.111, 2.815, 't'], [5.814, 2.815, 'b'], [7.519, 2.815, 't']]
        .forEach(([cx, cy, side]) => arcBand(s, cx, cy, 0.916, 0.623, 0.07,
            side === 'b' ? 0 : 180, side === 'b' ? 180 : 360, SHADE));

    const nodes = [
        { x: 0.820, c: NAVY,  year: '2020' },
        { x: 2.523, c: TEAL,  year: '2021' },
        { x: 4.226, c: MINT,  year: '2022' },
        { x: 5.929, c: GREEN, year: '2023' },
        { x: 7.631, c: LEAF,  year: '2024' },
    ];
    nodes.forEach(n => {
        shp(s, 'ellipse', n.x, 2.082, 1.466, 1.466, { fill: F(n.c) });
        dot(s, n.x + 0.733, 2.815, 1.219, WHITE);
        txt(s, n.year, { x: n.x + 0.209, y: 2.55, w: 1.049, h: 0.53, fontSize: 27, bold: true, color: n.c, align: 'center' });
    });

    [{ x: 1.962, y: 3.548, c: NAVY,  name: 'Timeline One',   tx: 1.576, ty: 4.589 },
     { x: 3.643, y: 1.191, c: TEAL,  name: 'Timeline Two',   tx: 3.281, ty: 0.266 },
     { x: 5.392, y: 3.548, c: MINT,  name: 'Timeline Three', tx: 4.984, ty: 4.589 },
     { x: 7.074, y: 1.191, c: GREEN, name: 'Timeline Four',  tx: 6.688, ty: 0.266 }].forEach(b => {
        shp(s, 'ellipse', b.x, b.y, 0.89, 0.89, { fill: F(b.c) });
        icon(s, b.x + 0.445, b.y + 0.445, 0.6, WHITE);
        txt(s, b.name, { x: b.tx, y: b.ty, w: 1.662, h: 0.303, fontSize: 14, bold: true, color: b.c, align: 'center' });
        body(s, L_SHORT, b.tx, b.ty + 0.303, 1.662, 0.465);
    });
}

/* =================================================== 5. PRESENTATION ===== */
const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
                slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
                slide21, slide22, slide23, slide24, slide25, slide26, slide27, slide28, slide29, slide30];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'TIMESKIP', width: 10, height: 5.625 });
pptx.layout = 'TIMESKIP';
pptx.author = 'Timeskip';
pptx.title = 'Timeline Infographic Presentation Template';
SLIDES.forEach(fn => { const s = pptx.addSlide(); s.background = { color: WHITE }; fn(s); });
pptx.writeFile({ fileName: path.join(__dirname, OUTPUT) })
    .then(f => console.log('wrote ' + f))
    .catch(e => { console.error(e); process.exit(1); });
