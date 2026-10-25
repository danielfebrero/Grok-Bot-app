/*
 * "Pinnacle - Digital Payment Presentation" - 40 slide deck rebuilt with pptxgenjs.
 *
 * Raster artwork from the original deck is replaced by flat "[image]" placeholder
 * rectangles; icon artwork is redrawn from primitive shapes.
 *
 *   node 122f2088-c2ca-452e-9592-2fa816608a8a_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const C = {
    pink: 'C647A0',
    blue: '143CB5',
    violet: '483FAF',
    indigo: '4A3FAF',
    yellow: 'FFC000',
    yellowDk: 'BF9000',
    ink: '404040',
    ink2: '595959',
    body: '808080',
    mute: 'BFBFBF',
    hair: 'D9D9D9',
    pale: 'F2F2F2',
    white: 'FFFFFF',
    grad1: '6C41AB',  // flat stand-in for the pink->blue gradient used on numerals
    grad2: '4A3FAF',
    placeholder: 'E4E1EA'
};

const F = {
    med: 'Poppins Medium',
    semi: 'Poppins SemiBold',
    reg: 'Poppins',
    light: 'Poppins Light',
    body: 'Lato',
    display: 'Raleway Black',
    displaySemi: 'Raleway SemiBold',
    quote: 'Lato Black'
};

/**
 * Every gradient in the deck is a 45-degree ramp built from the same three
 * brand colours.  Named ramps, listed as [position, colour] stops.
 */
const G = {
    main: [[0, C.pink], [1, C.blue]],            // full pink -> blue
    card: [[0, '6041AC'], [1, C.blue]],          // violet -> blue (cards, chips, tabs)
    band: [[0, C.pink], [1, '6041AC']],          // pink -> violet (banners, badges)
    deep: [[0, C.pink], [1, C.indigo]],          // pink -> indigo (full-height panels)
    steel: [[0, C.pink], [1, '353EB1']],         // pink -> steel blue (side bars)
    hero: [[0, '7762B8'], [1, '1A3EB6']],        // muted violet -> blue (slide 16 banner)
    panel: [[0, '6A4DB0'], [1, '173CB4']],       // slide 21 translucent panel
    quote: [[0, '664BB0'], [1, '294DBA']],       // slide 39 translucent quote block
    swot: [[0, 'CF6AB1'], [1, '4E57BD']],       // translucent SWOT wedge
    tile: [[0, 'C74CA2'], [1, 'A152AC']]         // slide 31 wallpaper
};

const SHADOW = { type: 'outer', color: '808080', opacity: 0.16, blur: 6, offset: 3, angle: 90 };

/* ------------------------------------------------------------------ helpers */

const hex2rgb = (h) => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
const rgb2hex = (c) => c.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0').toUpperCase()).join('');

/** colour at position t (0..1) along a stop list */
function stopColor(stops, t) {
    if (t <= stops[0][0]) return stops[0][1];
    for (let i = 1; i < stops.length; i++) {
        if (t <= stops[i][0]) {
            const [p0, c0] = stops[i - 1];
            const [p1, c1] = stops[i];
            const k = (t - p0) / (p1 - p0 || 1);
            const a = hex2rgb(c0);
            const b = hex2rgb(c1);
            return rgb2hex(a.map((v, j) => v + (b[j] - v) * k));
        }
    }
    return stops[stops.length - 1][1];
}

/**
 * pptxgenjs has no gradient fill, so linear gradients are painted as a grid of
 * flat tiles.  `ang` is the OOXML angle (45 = down/right, 90 = down, ...).
 */
function gradRect(sl, x, y, w, h, stops, ang, steps) {
    const rad = (ang === undefined ? 45 : ang) * Math.PI / 180;
    const cx = Math.cos(rad);
    const cy = Math.sin(rad);
    const span = Math.abs(cx) * w + Math.abs(cy) * h;
    const n = steps || 26;
    const nu = Math.max(1, Math.round(n * Math.abs(cx) * w / span));
    const nv = Math.max(1, Math.round(n * Math.abs(cy) * h / span));
    const ov = 0.012;
    for (let i = 0; i < nu; i++) {
        for (let j = 0; j < nv; j++) {
            const px = ((i + 0.5) / nu) * w;
            const py = ((j + 0.5) / nv) * h;
            const t = ((cx >= 0 ? px : w - px) * Math.abs(cx) + (cy >= 0 ? py : h - py) * Math.abs(cy)) / span;
            sl.addShape('rect', {
                x: x + i * w / nu - (i ? ov : 0), y: y + j * h / nv - (j ? ov : 0),
                w: w / nu + (i ? ov : 0), h: h / nv + (j ? ov : 0),
                fill: { color: stopColor(stops, t) }, line: { type: 'none' }
            });
        }
    }
}

/**
 * Background-coloured wedge that rounds off one corner of a painted gradient:
 * the corner square minus the quarter disc that the rounded shape keeps.
 */
function corner(sl, x, y, r, which, bg) {
    const K = 0.5523 * r;                              // circle -> bezier constant
    const cnr = [which.indexOf('l') >= 0 ? 0 : r, which.indexOf('t') >= 0 ? 0 : r];
    const hub = [r - cnr[0], r - cnr[1]];              // centre of the quarter disc
    const a = [hub[0], cnr[1]];                        // arc start (vertically level with corner)
    const b = [cnr[0], hub[1]];                        // arc end
    sl.addShape('custGeom', {
        x: x - cnr[0], y: y - cnr[1], w: r, h: r,
        points: [
            { x: cnr[0], y: cnr[1], moveTo: true },
            { x: a[0], y: a[1] },
            {
                x: b[0], y: b[1],
                curve: {
                    type: 'cubic',
                    x1: a[0] + (cnr[0] - hub[0]) / r * K, y1: a[1],
                    x2: b[0], y2: b[1] + (cnr[1] - hub[1]) / r * K
                }
            },
            { close: true }
        ],
        fill: { color: bg || C.white }, line: { type: 'none' }
    });
}

/** gradient rounded rectangle = gradient tiles + background wedges on the corners */
function gradRound(sl, x, y, w, h, r, stops, ang, corners, bg) {
    gradRect(sl, x, y, w, h, stops, ang);
    (corners || ['tl', 'tr', 'bl', 'br']).forEach((k) => {
        corner(sl, k.indexOf('l') >= 0 ? x : x + w, k.indexOf('t') >= 0 ? y : y + h, r, k, bg);
    });
}

/** annular sector, optionally colour-graded along the sweep */
function ring(sl, cx, cy, rOut, thick, startAng, sweep, color, segs) {
    const n = segs || 1;
    const rIn = rOut - thick;
    const rad = (d) => d * Math.PI / 180;
    for (let s = 0; s < n; s++) {
        const a0 = startAng + sweep * s / n;
        const a1 = startAng + sweep * (s + 1) / n;
        const sw = a1 - a0 + (n > 1 ? Math.sign(sweep) * 0.6 : 0);
        const col = typeof color === 'function' ? color((s + 0.5) / n) : color;
        sl.addShape('custGeom', {
            x: cx - rOut, y: cy - rOut, w: 2 * rOut, h: 2 * rOut,
            points: [
                { x: rOut + rOut * Math.cos(rad(a0)), y: rOut + rOut * Math.sin(rad(a0)), moveTo: true },
                {
                    x: rOut + rOut * Math.cos(rad(a0 + sw)), y: rOut + rOut * Math.sin(rad(a0 + sw)),
                    curve: { type: 'arc', hR: rOut, wR: rOut, stAng: a0, swAng: sw }
                },
                { x: rOut + rIn * Math.cos(rad(a0 + sw)), y: rOut + rIn * Math.sin(rad(a0 + sw)) },
                {
                    x: rOut + rIn * Math.cos(rad(a0)), y: rOut + rIn * Math.sin(rad(a0)),
                    curve: { type: 'arc', hR: rIn, wR: rIn, stAng: a0 + sw, swAng: -sw }
                },
                { close: true }
            ],
            fill: { color: col }, line: { type: 'none' }
        });
    }
}

/* -------------------------------------------------------------- text styles */

const base = (o) => Object.assign({ valign: 'top', isTextBox: true }, o);

/** small magenta section label, e.g. "About Us" */
function kicker(sl, x, y, t, o) {
    sl.addText(t, base(Object.assign({ x, y, w: 2.2, h: 0.37, fontFace: F.med, fontSize: 16, color: C.pink, wrap: false }, o)));
}
/** big dark headline */
function title(sl, x, y, w, h, t, o) {
    sl.addText(t, base(Object.assign({ x, y, w, h, fontFace: F.semi, fontSize: 28, color: C.ink }, o)));
}
/** grey running text, justified with 1.5 line spacing */
function body(sl, x, y, w, h, t, o) {
    sl.addText(t, base(Object.assign({
        x, y, w, h, fontFace: F.body, fontSize: 11, color: C.body,
        align: 'justify', lineSpacingMultiple: 1.5
    }, o)));
}
/** mid-grey sub-heading */
function label(sl, x, y, w, t, o) {
    sl.addText(t, base(Object.assign({ x, y, w, h: 0.34, fontFace: F.med, fontSize: 14, color: C.ink2 }, o)));
}
/** big gradient-coloured statistic */
function stat(sl, x, y, w, h, t, size, color) {
    sl.addText(t, base({ x, y, w, h, fontFace: F.med, fontSize: size, color: color || C.grad1, align: 'center', wrap: false }));
}
/** pill shaped call-to-action */
function pill(sl, x, y, w, h, t, o) {
    o = o || {};
    sl.addShape('roundRect', {
        x, y, w, h, rectRadius: h / 2, fill: { color: C.pink }, line: { type: 'none' }
    });
    gradRound(sl, x, y, w, h, h / 2, G.main, 45, ['tl', 'tr', 'bl', 'br'], o.bg || C.white);
    sl.addText(t, {
        x, y, w, h, fontFace: o.font || F.reg, fontSize: o.size || 11, color: C.pale,
        align: 'center', valign: 'middle', margin: 0
    });
}

/* ------------------------------------------------------------ icon vignettes */

/** outline "card in hand" mark */
function iconCard(sl, x, y, s, color) {
    const c = color || C.violet;
    const lw = Math.max(0.75, s * 2.2);
    sl.addShape('roundRect', { x, y, w: s * 0.86, h: s * 0.58, rectRadius: s * 0.06, rotate: -12, fill: { type: 'none' }, line: { color: c, width: lw } });
    sl.addShape('rect', { x: x + s * 0.03, y: y + s * 0.14, w: s * 0.82, h: s * 0.09, rotate: -12, fill: { color: c }, line: { type: 'none' } });
    sl.addShape('rect', { x: x + s * 0.1, y: y + s * 0.36, w: s * 0.3, h: s * 0.05, rotate: -12, fill: { color: c }, line: { type: 'none' } });
    // thumb reaching over the lower right corner
    sl.addShape('blockArc', { x: x + s * 0.46, y: y + s * 0.3, w: s * 0.6, h: s * 0.72, rotate: 210, fill: { type: 'none' }, line: { color: c, width: lw } });
}
/** outline "smartphone paying" mark */
function iconPhone(sl, x, y, s, color) {
    const c = color || C.violet;
    const lw = Math.max(0.75, s * 2.2);
    sl.addShape('roundRect', { x: x + s * 0.22, y: y + s * 0.1, w: s * 0.56, h: s * 0.9, rectRadius: s * 0.08, fill: { type: 'none' }, line: { color: c, width: lw } });
    sl.addShape('roundRect', { x: x + s * 0.06, y, w: s * 0.62, h: s * 0.42, rectRadius: s * 0.05, fill: { color: C.white }, line: { color: c, width: lw } });
    sl.addShape('rect', { x: x + s * 0.06, y: y + s * 0.1, w: s * 0.62, h: s * 0.08, fill: { color: c }, line: { type: 'none' } });
    sl.addShape('ellipse', { x: x + s * 0.45, y: y + s * 0.88, w: s * 0.1, h: s * 0.08, fill: { type: 'none' }, line: { color: c, width: lw } });
}
/** stacked gold bars mark */
function iconBars(sl, x, y, s, color) {
    const c = color || C.violet;
    const lw = Math.max(0.75, s * 1.7);
    const bw = s * 0.32;
    const bh = s * 0.23;
    [[0, 1], [1, 0.5], [1, 1.5], [2, 0], [2, 1], [2, 2]].forEach(([row, col]) => {
        sl.addShape('trapezoid', {
            x: x + col * bw + (s - 3 * bw) / 2, y: y + s * 0.28 + row * bh, w: bw * 0.96, h: bh * 0.95,
            fill: { type: 'none' }, line: { color: c, width: lw }
        });
    });
    sl.addShape('star4', { x: x + s * 0.06, y: y + s * 0.02, w: s * 0.19, h: s * 0.19, fill: { color: c }, line: { type: 'none' } });
}
/** coin resting on an open hand */
function iconCoinHand(sl, x, y, s, color) {
    const c = color || C.violet;
    const lw = Math.max(0.75, s * 2.0);
    sl.addShape('ellipse', { x: x + s * 0.28, y: y + s * 0.05, w: s * 0.42, h: s * 0.42, fill: { type: 'none' }, line: { color: c, width: lw } });
    sl.addText('$', { x: x + s * 0.28, y: y + s * 0.02, w: s * 0.42, h: s * 0.42, fontFace: F.semi, fontSize: s * 24, color: c, align: 'center', valign: 'middle', margin: 0 });
    sl.addShape('star4', { x: x + s * 0.08, y: y + s * 0.1, w: s * 0.14, h: s * 0.14, fill: { color: c }, line: { type: 'none' } });
    sl.addShape('star4', { x: x + s * 0.76, y: y + s * 0.3, w: s * 0.13, h: s * 0.13, fill: { color: c }, line: { type: 'none' } });
    // open palm underneath
    sl.addShape('blockArc', { x, y: y + s * 0.35, w: s, h: s * 0.85, rotate: 180, fill: { type: 'none' }, line: { color: c, width: lw } });
}
/** head-and-shoulders silhouette */
function iconUser(sl, x, y, s, color) {
    const c = color || C.white;
    sl.addShape('ellipse', { x: x + s * 0.3, y, w: s * 0.4, h: s * 0.4, fill: { color: c }, line: { type: 'none' } });
    sl.addShape('blockArc', { x, y: y + s * 0.42, w: s, h: s * 0.95, fill: { color: c }, line: { type: 'none' } });
}
/** tick mark drawn from two rotated bars */
function iconCheck(sl, cx, cy, s, color) {
    const c = color || C.white;
    sl.addShape('rect', { x: cx - s * 0.42, y: cy - s * 0.02, w: s * 0.38, h: s * 0.15, rotate: 45, fill: { color: c }, line: { type: 'none' } });
    sl.addShape('rect', { x: cx - s * 0.18, y: cy - s * 0.06, w: s * 0.66, h: s * 0.15, rotate: -45, fill: { color: c }, line: { type: 'none' } });
}
/** map pin */
function iconPin(sl, x, y, s, color) {
    sl.addShape('teardrop', { x, y, w: s, h: s, rotate: 225, fill: { color: color || C.pink }, line: { type: 'none' } });
    sl.addShape('ellipse', { x: x + s * 0.34, y: y + s * 0.26, w: s * 0.3, h: s * 0.3, fill: { color: C.white }, line: { type: 'none' } });
}
/** handset */
function iconReceiver(sl, x, y, s, color) {
    sl.addShape('blockArc', { x, y, w: s, h: s, rotate: 200, fill: { color: color || C.pink }, line: { type: 'none' } });
}
/**
 * Photo frame inherited from the template layouts.  Every one of them ships
 * empty in the source deck, so the frame is kept (it documents the layout) but
 * drawn unfilled - exactly how the original renders.
 */
function photoSlot(sl, x, y, w, h, o) {
    o = o || {};
    sl.addShape(o.shape || 'rect', Object.assign({
        x, y, w, h, fill: { type: 'none' }, line: { type: 'none' }
    }, o.rectRadius ? { rectRadius: o.rectRadius } : {}));
}

/** grey box + caption standing in for a raster image of the original deck */
function imagePlaceholder(sl, x, y, w, h, o) {
    o = o || {};
    const geo = Object.assign({ x, y, w, h }, o.rotate ? { rotate: o.rotate } : {});
    sl.addShape(o.shape || 'rect', Object.assign({}, geo, {
        fill: { color: o.color || C.placeholder }, line: { type: 'none' }
    }, o.rectRadius ? { rectRadius: o.rectRadius } : {}));
    sl.addText('[image]', Object.assign({}, geo, {
        fontFace: F.med, fontSize: 11, color: '9A94A6', align: 'center', valign: 'middle', margin: 0
    }));
}

/* ------------------------------------------------------------- deck furniture */

const RAIL_X = 12.3214;

/** thin vertical hairline that separates the side rail on most slides */
function rail(sl) {
    sl.addShape('line', { x: RAIL_X, y: 0, w: 0, h: 7.5, line: { color: C.pale, width: 3 } });
}

/** coin logo used in the rail and, larger, on the cover */
function coin(sl, x, y, s) {
    sl.addShape('ellipse', { x, y, w: s, h: s, fill: { color: C.yellow }, line: { type: 'none' } });
    const p = s * 0.155;
    const d = s - 2 * p;
    gradRound(sl, x + p, y + p, d, d, d / 2, G.main, 45, ['tl', 'tr', 'bl', 'br'], C.yellow);
    ring(sl, x + s / 2, y + s / 2, s * 0.325, s * 0.055, 100, 150, C.pale);
    ring(sl, x + s / 2, y + s / 2, s * 0.325, s * 0.055, -80, 150, C.pale);
    sl.addText('$', { x, y: y - s * 0.03, w: s, h: s, fontFace: F.semi, fontSize: s * 62, color: C.pale, align: 'center', valign: 'middle', margin: 0 });
}

/** rail widgets: logo badge, four progress dots and the page number */
const DOTS = [[12.608, 2.836, 0.439], [12.642, 3.425, 0.371], [12.664, 3.946, 0.326], [12.706, 4.422, 0.242]];

function chrome(sl, num, activeDot) {
    rail(sl);
    sl.addShape('ellipse', { x: 12.5273, y: 0.2821, w: 0.5774, h: 0.5783, fill: { color: C.pale }, line: { type: 'none' } });
    coin(sl, 12.5968, 0.3517, 0.4388);
    DOTS.forEach(([x, y, d], i) => {
        if (i === activeDot) gradRound(sl, x, y, d, d, d / 2, G.main, 45, ['tl', 'tr', 'bl', 'br'], C.white);
        else sl.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: C.hair, width: 0.75 } });
    });
    sl.addText(String(num).padStart(2, '0'), {
        x: 12.5576, y: 6.806, w: 0.5395, h: 0.3118, fontFace: F.semi, fontSize: 16,
        color: '8B8B8B', align: 'center', valign: 'middle', margin: 0
    });
}

/* ------------------------------------------------------------------- slides */

const LOREM_LONG = "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  " +
    "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  " +
    'when an unknown printer took a galley of type and scrambled it to make a type specimen book.';
const LOREM_MED = "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  " +
    "Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  " +
    'when an unknown printer took a galley.';
const LOREM_SHORT = "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  " +
    "Lorem Ipsum has been the industry's standard dummy text. ";
const LOREM_TINY = "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  " +
    "Lorem Ipsum has been the industry's.";

const builders = [];
const slide = (fn) => builders.push(fn);

/* 1 - cover */
slide((sl) => {
    sl.addShape('rect', { x: 1.6696, y: 0, w: 11.6638, h: 7.5, fill: { color: C.pale }, line: { type: 'none' } });
    gradRect(sl, 0, 1.0714, 12.9718, 5.3571, G.main, 45, 40);
    sl.addShape('rect', { x: 1.6669, y: 0.3257, w: 11.6422, h: 0.377, fill: { color: C.white }, line: { type: 'none' } });
    photoSlot(sl, 0, 1.281, 8.375, 4.9381);

    [['PAYMENT', 2.8412, 0.959], ['WALLET', 5.5652, 0.812], ['FINANCE', 8.52, 0.898], ['MARKETING', 11.0428, 1.133]]
        .forEach(([t, x, w]) => sl.addText(t, base({
            x, y: 0.3628, w, h: 0.3029, fontFace: F.med, fontSize: 12, color: C.ink, align: 'center', margin: 0
        })));

    coin(sl, 10.132, 2.1942, 0.9256);
    sl.addShape('star4', { x: 9.8829, y: 1.9089, w: 0.2848, h: 0.2853, fill: { color: C.white }, line: { type: 'none' } });
    sl.addShape('star4', { x: 11.1645, y: 2.8004, w: 0.2848, h: 0.2853, fill: { color: C.white }, line: { type: 'none' } });

    sl.addText('Pinnacle', base({ x: 9.061, y: 3.3073, w: 3.2102, h: 0.9424, fontFace: F.display, fontSize: 50, color: C.white, align: 'center', margin: 0 }));
    sl.addText('Digital Payment Presentation', base({ x: 9.3143, y: 4.1908, w: 2.7036, h: 0.3029, fontFace: F.med, fontSize: 12, color: C.hair, align: 'center', margin: 0 }));
    sl.addText('Presented by\nNew Coral Std.', base({ x: 8.4738, y: 5.7413, w: 1.3134, h: 0.4712, fontFace: F.med, fontSize: 11, color: C.pale, margin: 0 }));
    body(sl, 2.8412, 6.5819, 9.2967, 0.62, LOREM_LONG + '  ');
});

/* 2 - welcome message */
slide((sl) => {
    chrome(sl, 2, 0);
    photoSlot(sl, 0, 0.6932, 4.4284, 6.1135);
    sl.addShape('roundRect', { x: 3.8333, y: 3.0179, w: 7.9711, h: 3.4237, rectRadius: 0.226, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });

    kicker(sl, 5.0262, 0.7587, 'About Us');
    title(sl, 5.0071, 1.1237, 4.1975, 0.5722, 'Welcome Message');
    body(sl, 5.0281, 1.9147, 5.9248, 0.8975, LOREM_MED);

    photoSlot(sl, 4.58, 3.5, 1.715, 1.6518, { shape: 'ellipse' });
    sl.addText('Michael Bennington', base({ x: 4.439, y: 5.2618, w: 2.0747, h: 0.3198, fontFace: F.med, fontSize: 13, color: C.ink2, align: 'center' }));
    sl.addText('Founder & CEO', base({ x: 4.4358, y: 5.5654, w: 2.0747, h: 0.2861, fontFace: F.reg, fontSize: 11, italic: true, color: C.body, align: 'center' }));
    sl.addText('\u201C', base({ x: 6.7068, y: 3.8981, w: 0.6782, h: 1.4473, fontFace: F.semi, fontSize: 80, color: C.yellow }));
    body(sl, 6.7482, 4.5002, 4.4304, 1.1752, LOREM_MED);
});

/* 3 - mission / vision */
slide((sl) => {
    chrome(sl, 3, 1);
    photoSlot(sl, 1.9532, 0, 3.1142, 3.75);
    photoSlot(sl, 6.0287, 3.7457, 3.1142, 3.7529);

    kicker(sl, 0.8957, 4.6023, 'About Us');
    title(sl, 0.8766, 4.9672, 4.4192, 1.5146, 'Solutions Platform for Money in One Place');

    label(sl, 5.9878, 0.8726, 1.7, 'Our Mission ', { fontSize: 18, h: 0.4 });
    label(sl, 9.8757, 4.6418, 1.7, 'Our Vision ', { fontSize: 18, h: 0.4 });
    [1.4663, 2.2176].forEach((y, i) => {
        body(sl, 6.5089, y, 4.5596, 0.62, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy.");
        sl.addShape('ellipse', { x: 6.1006, y: 1.5992 + i * 0.7641, w: 0.1779, h: 0.1782, fill: { color: C.yellow }, line: { type: 'none' } });
        iconCheck(sl, 6.1895, 1.688 + i * 0.7641, 0.14);
    });
    body(sl, 9.8757, 5.2293, 2.0112, 1.1752,
        '\u201CLorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has.\u201D', { italic: true });
});

/* 4 - two stat cards */
slide((sl) => {
    chrome(sl, 4, 2);
    photoSlot(sl, 7.362, 0, 4.4431, 6.6957);

    kicker(sl, 0.896, 1.774, 'About Us');
    title(sl, 0.877, 2.139, 3.368, 1.986, 'Focus to Help Your Business with Digital Payment');
    body(sl, 0.898, 4.273, 3.368, 1.453, LOREM_MED);

    // filled card
    gradRound(sl, 5.646, 2.813, 2.375, 3.392, 0.198, G.card, 45);
    sl.addShape('ellipse', { x: 6.105, y: 3.285, w: 1.458, h: 1.447, fill: { type: 'none' }, line: { color: C.white, width: 2.5 } });
    sl.addShape('ellipse', { x: 6.509, y: 4.434, w: 0.649, h: 0.645, fill: { color: C.violet }, line: { type: 'none' } });
    sl.addShape('heart', { x: 6.684, y: 4.61, w: 0.297, h: 0.292, fill: { color: C.white }, line: { type: 'none' } });
    stat(sl, 6.269, 3.737, 1.155, 0.555, '300+', 27, C.white);
    sl.addText('Happy Customers & Companies', base({ x: 6.019, y: 5.262, w: 1.629, h: 0.471, fontFace: F.med, fontSize: 11, color: C.white, align: 'center' }));

    // white card
    sl.addShape('roundRect', { x: 8.295, y: 2.813, w: 2.375, h: 3.392, rectRadius: 0.198, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
    sl.addShape('ellipse', { x: 8.754, y: 3.285, w: 1.458, h: 1.447, fill: { type: 'none' }, line: { color: C.pink, width: 2.5 } });
    sl.addShape('ellipse', { x: 9.158, y: 4.434, w: 0.649, h: 0.645, fill: { color: C.white }, line: { type: 'none' } });
    iconUser(sl, 9.346, 4.61, 0.29, C.ink);
    stat(sl, 8.953, 3.737, 1.061, 0.555, '150+', 27);
    sl.addText('Best Quality Consultants', base({ x: 8.882, y: 5.262, w: 1.169, h: 0.471, fontFace: F.med, fontSize: 11, color: C.ink2, align: 'center' }));
});

/* 5 - "180% sales growth" card */
slide((sl) => {
    chrome(sl, 5, 3);
    sl.addShape('roundRect', { x: 1.6867, y: 1.6571, w: 4.9094, h: 4.7845, rectRadius: 0.28, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
    photoSlot(sl, 1.0246, 3.2292, 2.4751, 3.6879, { shape: 'roundRect', rectRadius: 0.126 });
    photoSlot(sl, 4.7803, 0.93, 3.0735, 3.5495, { shape: 'roundRect', rectRadius: 0.156 });

    sl.addShape('roundRect', { x: 3.227, y: 2.559, w: 2.48, h: 2.98, rectRadius: 0.145, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
    stat(sl, 3.795, 2.84, 1.343, 0.707, '180%', 36);
    sl.addText('Sales Growth', base({ x: 3.545, y: 3.511, w: 1.843, h: 0.286, fontFace: F.med, fontSize: 11, color: C.ink2, align: 'center' }));
    body(sl, 3.652, 3.781, 1.63, 0.897, 'Lorem Ipsum  simply\n dummy text  printing\n typesetting industry.', { align: 'center' });
    // gradient footer band with rounded bottom corners
    gradRound(sl, 3.227, 4.959, 2.48, 0.58, 0.145, G.main, 45, ['bl', 'br']);
    iconCheck(sl, 3.85, 5.248, 0.28);
    sl.addText('Completed', base({ x: 4.122, y: 5.106, w: 1.116, h: 0.286, fontFace: F.med, fontSize: 11, color: C.white, align: 'center' }));

    kicker(sl, 8.572, 1.918, 'About Us');
    title(sl, 8.553, 2.283, 3.397, 1.043, 'Make Payments Very Easily');
    body(sl, 8.573, 3.551, 3.315, 1.453, LOREM_MED);
    pill(sl, 8.66, 5.229, 1.427, 0.334, 'Learn More');
});

/* 6 - digital currency banner */
slide((sl) => {
    chrome(sl, 6, 0);
    photoSlot(sl, 0, 3.7534, 11.8083, 3.7466);
    gradRound(sl, 6.4342, 0.7402, 5.3849, 1.2303, 0.072, G.band, 45);

    kicker(sl, 0.896, 0.742, 'About Us');
    title(sl, 0.877, 1.107, 5.003, 1.043, 'Significance of A Digital Wallets');
    body(sl, 0.898, 2.286, 4.569, 0.897,
        "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer");

    sl.addShape('ellipse', { x: 6.236, y: 1.152, w: 0.398, h: 0.407, fill: { color: C.yellow }, line: { type: 'none' } });
    sl.addText('\u203A', { x: 6.236, y: 1.115, w: 0.398, h: 0.407, fontFace: F.semi, fontSize: 18, color: C.white, align: 'center', valign: 'middle', margin: 0 });

    label(sl, 7.042, 0.89, 2.522, 'Digital Currency', { color: C.pale });
    body(sl, 7.048, 1.201, 4.231, 0.62, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been.', { color: C.hair });
    label(sl, 7.042, 2.208, 2.522, 'Mobile Wallet', { color: C.mute });
    body(sl, 7.048, 2.52, 4.231, 0.62, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been.', { color: C.mute });
});

/* 7 - era of digital payment */
slide((sl) => {
    chrome(sl, 7, 1);
    [[1.5376, 0.3105, 2.42, 2.3309], [4.6672, 2.4901, 2.1298, 2.0514], [1.0262, 3.75, 3.2619, 3.1418]]
        .forEach(([x, y, w, h]) => photoSlot(sl, x, y, w, h, { shape: 'ellipse' }));

    stat(sl, 4.726, 0.905, 1.158, 0.555, '450+', 27);
    body(sl, 4.751, 1.427, 2.412, 0.62, 'Lorem Ipsum is simply dummy text of the printing.');
    sl.addShape('ellipse', { x: 4.484, y: 1.576, w: 0.145, h: 0.145, fill: { color: C.yellow }, line: { type: 'none' } });
    stat(sl, 4.796, 5.607, 1.061, 0.555, '150+', 27);
    body(sl, 4.821, 6.13, 2.412, 0.62, 'Lorem Ipsum is simply dummy text of the printing.');
    sl.addShape('ellipse', { x: 4.554, y: 6.279, w: 0.145, h: 0.145, fill: { color: C.yellow }, line: { type: 'none' } });

    kicker(sl, 8.538, 1.605, 'About Us');
    title(sl, 8.519, 1.97, 3.397, 1.043, 'It\u2019s the Era of Digital Payment');
    body(sl, 8.533, 3.249, 3.382, 1.453, LOREM_LONG.replace('it to make a type specimen book.', '') + ' ');
    sl.addText('226', base({ x: 8.513, y: 5.015, w: 1.158, h: 0.555, fontFace: F.med, fontSize: 27, color: C.grad1 }));
    sl.addText('1283', base({ x: 10.546, y: 5.015, w: 1.158, h: 0.555, fontFace: F.med, fontSize: 27, color: C.grad1 }));
    sl.addText('Branch Office', base({ x: 8.522, y: 5.592, w: 1.395, h: 0.303, fontFace: F.med, fontSize: 12, color: C.ink2 }));
    sl.addText('Partner Global', base({ x: 10.523, y: 5.592, w: 1.438, h: 0.303, fontFace: F.med, fontSize: 12, color: C.ink2 }));
});

/* 8 - numbered two column */
slide((sl) => {
    chrome(sl, 8, 2);
    photoSlot(sl, 0, 0, 6.2451, 3.26);
    photoSlot(sl, 9.2072, 3.8121, 2.5925, 3.6879);

    kicker(sl, 7.248, 0.886, 'About Us');
    title(sl, 7.228, 1.251, 3.958, 1.515, 'High & Certified Quality for Finance Business');

    [['01.', 'Grown As Digital', 0.875, 0.898, 2.158], ['02.', 'Big Profit Company', 5.006, 5.029, 6.289]]
        .forEach(([num, head, nx, tx, hx]) => {
            sl.addText(num, base({ x: nx, y: 3.986, w: 1.322, h: 1.01, fontFace: F.semi, fontSize: 54, bold: true, color: C.grad1 }));
            sl.addText(head, base({ x: hx, y: 4.274, w: 2.075, h: 0.286, fontFace: F.med, fontSize: 11, color: C.ink2 }));
            sl.addText('About Us', base({ x: hx - 0.003, y: 4.555, w: 2.075, h: 0.269, fontFace: F.reg, fontSize: 10, color: C.body }));
            body(sl, tx, 5.051, 3.468, 1.731, LOREM_LONG);
        });
});

/* 9 - three numbered circles */
slide((sl) => {
    chrome(sl, 9, 3);
    [1.0, 4.9774, 8.8803].forEach((x) => photoSlot(sl, x, 3.25, 2.9266, 2.7605));

    kicker(sl, 0.896, 1.215, 'About Us');
    title(sl, 0.877, 1.58, 5.003, 0.976, 'Transaction That Takes Place Via Digital', { fontSize: 26 });
    body(sl, 6.689, 1.298, 5.232, 1.175, LOREM_LONG);

    [[1.974, '01.'], [5.951, '02.'], [9.854, '03.']].forEach(([x, t]) => {
        gradRound(sl, x, 5.513, 0.979, 0.979, 0.49, G.card, 45);
        sl.addShape('ellipse', { x, y: 5.513, w: 0.979, h: 0.979, fill: { type: 'none' }, line: { color: C.white, width: 2.5 } });
        sl.addText(t, base({ x: x + 0.084, y: 5.75, w: 0.81, h: 0.505, fontFace: F.semi, fontSize: 24, bold: true, color: C.white, align: 'center' }));
    });
});

/* 10 - board of professional */
slide((sl) => {
    chrome(sl, 10, 0);
    gradRound(sl, 7.3165, 1.9077, 4.4905, 3.7673, 0.255, G.card, 45);
    photoSlot(sl, 1.0104, 2.3526, 3.5153, 2.7605, { shape: 'roundRect', rectRadius: 0.252 });
    photoSlot(sl, 4.8605, 2.3526, 3.5153, 2.7605, { shape: 'roundRect', rectRadius: 0.252 });

    kicker(sl, 0.896, 1.167, 'Our Team');
    title(sl, 0.877, 1.532, 5.003, 0.538, 'Board of Professional', { fontSize: 26 });

    [['Danilla Eduardo', 'Co-Founder', 1.111], ['Evangeline Shaw', 'Senior Director', 4.911]].forEach(([n, r, x]) => {
        sl.addText(n, base({ x, y: 5.379, w: 2.437, h: 0.37, fontFace: F.med, fontSize: 16, color: C.ink2 }));
        sl.addText(r, base({ x: x - 0.004, y: 5.713, w: 2.437, h: 0.303, fontFace: F.reg, fontSize: 12, italic: true, color: C.body }));
    });

    label(sl, 8.746, 2.812, 2.249, 'What We Envision ?', { color: C.pale });
    body(sl, 8.731, 3.229, 2.588, 1.175, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard. ", { color: C.hair });
    sl.addShape('ellipse', { x: 10.82, y: 4.903, w: 0.398, h: 0.407, fill: { color: C.yellow }, line: { type: 'none' } });
    sl.addText('\u203A', { x: 10.82, y: 4.866, w: 0.398, h: 0.407, fontFace: F.semi, fontSize: 18, color: C.white, align: 'center', valign: 'middle', margin: 0 });
});

/* 11 - four member cards */
slide((sl) => {
    chrome(sl, 11, 1);
    const team = [
        ['Alvaro Reyes', 'Executive Officer', 1.0333, true],
        ['Mathilda Langevin', 'General Counsel', 3.7741, false],
        ['Priscilla Dawson', 'Investment Admin', 6.5787, false],
        ['Jeremy Chouette', 'Financial Planner', 9.3834, false]
    ];
    team.forEach(([name, role, x, hi]) => {
        if (hi) gradRound(sl, x, 2.9545, 2.4243, 2.5227, 0.16, G.card, 45);
        else sl.addShape('roundRect', { x, y: 2.9545, w: 2.4243, h: 2.5227, rectRadius: 0.066, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
        photoSlot(sl, x + 0.4805, 2.2624, 1.4634, 1.4096, { shape: 'ellipse' });
        sl.addText(name, base({ x: x - 0.004, y: 4.059, w: 2.437, h: 0.32, fontFace: F.med, fontSize: 13, color: hi ? C.white : C.ink2, align: 'center' }));
        sl.addText(role, base({ x: x - 0.008, y: 4.362, w: 2.437, h: 0.286, fontFace: F.reg, fontSize: 11, italic: true, color: hi ? C.pale : C.body, align: 'center' }));
        sl.addShape('line', { x: x + 0.788, y: 5.008, w: 0.846, h: 0, line: { color: hi ? C.pale : C.blue, width: 2.25 } });
    });

    kicker(sl, 5.73, 0.829, 'Our Team', { w: 1.312, align: 'center' });
    title(sl, 3.462, 1.193, 5.753, 0.572, 'Pinnacle Solid Team', { align: 'center' });
    body(sl, 1.905, 6.137, 8.959, 0.62, LOREM_LONG, { align: 'center' });
});

/* 12 - expert team stats */
slide((sl) => {
    chrome(sl, 12, 2);
    photoSlot(sl, 0, 0, 5.9792, 7.5);

    kicker(sl, 6.236, 1.352, 'Our Expert Team', { w: 2.1 });
    title(sl, 6.217, 1.724, 4.367, 0.572, 'Arthur Emmanuelle');
    body(sl, 6.231, 2.435, 4.745, 1.175, LOREM_LONG.replace('it to make a type specimen book.', '') + ' ');

    const cards = [
        ['+85%', 'Client Happy', 3.009, true],
        ['6790', 'Success Branding', 6.159, false],
        ['1400+', 'Completed Progress', 9.343, false]
    ];
    cards.forEach(([num, cap, x, hi]) => {
        if (hi) gradRound(sl, x, 4.271, 2.464, 2.224, 0.13, G.card, 45);
        else sl.addShape('roundRect', { x, y: 4.271, w: 2.464, h: 2.224, rectRadius: 0.13, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
        sl.addText(num, base({ x: x + 0.31, y: 4.676, w: 1.846, h: 0.774, fontFace: F.med, fontSize: 40, color: hi ? C.white : C.pink, align: 'center' }));
        sl.addText(cap, base({ x: x + 0.34, y: 5.452, w: 1.78, h: 0.639, fontFace: F.med, fontSize: 16, color: hi ? C.pale : C.ink2, align: 'center' }));
    });
});

/* 13 - business team carousel */
slide((sl) => {
    chrome(sl, 13, 3);
    const team = [['Frans Nicotra', 'Financial Director', 1.244], ['Laura Ferrero', 'Customer Service', 5.072], ['Austin Souza', 'Sales Manager', 8.916]];
    team.forEach(([name, role, x]) => {
        photoSlot(sl, x - 0.233, 2.322, 3.125, 3.339);
        // white body of the card with rounded bottom corners
        sl.addShape('roundRect', { x, y: 5.608, w: 2.657, h: 1.01, rectRadius: 0.06, fill: { color: C.white }, line: { type: 'none' } });
        sl.addShape('rect', { x, y: 5.608, w: 2.657, h: 0.5, fill: { color: C.white }, line: { type: 'none' } });
        // name plate: gradient strip with a gold avatar block on the left
        gradRect(sl, x, 4.833, 2.657, 0.775, G.card, 45, 14);
        sl.addShape('rect', { x, y: 4.833, w: 0.711, h: 0.774, fill: { color: C.yellow }, line: { type: 'none' } });
        iconUser(sl, x + 0.187, 5.033, 0.338);
        sl.addText(name, base({ x: x + 0.855, y: 4.928, w: 1.651, h: 0.32, fontFace: F.med, fontSize: 13, color: C.white, align: 'center' }));
        sl.addText(role, base({ x: x + 0.855, y: 5.231, w: 1.651, h: 0.286, fontFace: F.reg, fontSize: 11, italic: true, color: C.pale, align: 'center' }));
        body(sl, x + 0.18, 5.811, 2.299, 0.62, 'Lorem Ipsum simply dummy text of the printing.', { align: 'center' });
    });

    kicker(sl, 5.73, 0.829, 'Our Team', { w: 1.312, align: 'center' });
    title(sl, 3.462, 1.193, 5.753, 0.572, 'Our Business Team', { align: 'center' });

    [[0.829, '\u2039'], [11.629, '\u203A']].forEach(([x, ch]) => {
        sl.addShape('ellipse', { x, y: 3.788, w: 0.398, h: 0.407, fill: { color: C.yellow }, line: { type: 'none' } });
        sl.addText(ch, { x, y: 3.751, w: 0.398, h: 0.407, fontFace: F.semi, fontSize: 18, color: C.white, align: 'center', valign: 'middle', margin: 0 });
    });
});

/* 14 - service tabs */
slide((sl) => {
    chrome(sl, 14, 0);
    // tab strip: first tab filled, remaining three outlined
    const tabs = [['Standard Payment', 2.129, true], ['Debit/Credit Card', 4.33, false], ['Overlay Services', 6.532, false], ['Mobile Wallets', 8.728, false]];
    tabs.forEach(([t, x, hi], i) => {
        if (hi) gradRound(sl, x, 2.2615, 2.1995, 1.5098, 0.08, G.card, 45, ['tl', 'tr']);
        else sl.addShape('roundRect', { x, y: 2.2615, w: 2.1995, h: 1.5098, rectRadius: 0.05, fill: { color: C.white }, line: { color: C.pale, width: 1 }, shadow: SHADOW });
        const cx = x + 1.1;
        if (i === 0) iconCoinHand(sl, cx - 0.36, 2.45, 0.62, C.white);
        if (i === 1) iconCard(sl, cx - 0.32, 2.5, 0.64);
        if (i === 2) iconBars(sl, cx - 0.32, 2.48, 0.6);
        if (i === 3) iconPhone(sl, cx - 0.32, 2.46, 0.62);
        sl.addText(t, base({ x: cx - 0.978, y: 3.247, w: 1.955, h: 0.303, fontFace: F.med, fontSize: 12, color: hi ? C.white : C.ink2, align: 'center' }));
    });

    photoSlot(sl, 1.0263, 4.164, 5.2556, 2.5227, { shape: 'roundRect', rectRadius: 0.23 });
    sl.addShape('roundRect', { x: 6.7444, y: 4.164, w: 5.0647, h: 2.5227, rectRadius: 0.083, fill: { type: 'none' }, line: { color: C.pale, width: 1.5 } });
    label(sl, 7.213, 4.493, 3.19, 'Benefits of Digital Payments', { color: C.ink });
    body(sl, 7.205, 4.868, 4.143, 0.897, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.");
    pill(sl, 7.292, 6.024, 1.427, 0.334, 'Learn More');

    kicker(sl, 5.64, 0.829, 'Our Service', { w: 1.492, align: 'center' });
    title(sl, 3.462, 1.193, 5.753, 0.572, 'Let\u2019s Check Our Services', { align: 'center' });
});

/* 15 - three service tiles */
slide((sl) => {
    chrome(sl, 15, 1);
    photoSlot(sl, 1.0263, 4.1754, 2.9003, 2.2632, { shape: 'roundRect', rectRadius: 0.17 });
    photoSlot(sl, 4.9979, 0.8778, 3.0471, 3.8064, { shape: 'roundRect', rectRadius: 0.23 });
    photoSlot(sl, 8.9126, 3.3139, 2.9003, 3.1247, { shape: 'roundRect', rectRadius: 0.22 });

    [[0.581, 3.734, 'bars'], [4.541, 1.216, 'coin'], [8.591, 3.018, 'phone']].forEach(([x, y, kind]) => {
        sl.addShape('roundRect', { x, y, w: 0.903, h: 0.913, rectRadius: 0.16, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
        if (kind === 'bars') iconBars(sl, x + 0.16, y + 0.165, 0.583);
        if (kind === 'coin') iconCoinHand(sl, x + 0.09, y + 0.19, 0.72);
        if (kind === 'phone') iconPhone(sl, x + 0.16, y + 0.165, 0.583);
    });

    kicker(sl, 0.896, 1.426, 'Our Services', { w: 1.7 });
    title(sl, 0.877, 1.791, 3.489, 1.414, 'What Are Digital Payment Services ?', { fontSize: 26 });
    label(sl, 8.737, 1.353, 2.082, 'Overlay Services');
    body(sl, 8.737, 1.728, 3.19, 0.897, LOREM_TINY);
    label(sl, 4.941, 5.271, 2.082, 'Mobile Wallets');
    body(sl, 4.941, 5.646, 3.19, 0.897, LOREM_TINY);
});

/* 16 - gradient banner with three circles */
slide((sl) => {
    chrome(sl, 16, 2);
    gradRound(sl, 1.0263, -0.012, 10.781, 4.191, 0.877, G.hero, 45, ['bl', 'br']);

    kicker(sl, 5.64, 0.829, 'Our Service', { w: 1.492, align: 'center', color: C.hair });
    title(sl, 3.167, 1.193, 6.499, 1.515, 'Digital Payments are About  to See Even More Changes in the Coming Future', { align: 'center', color: C.white });

    const cols = [
        [2.192, 'Mobile Wallets', 1.31, 1.864, 'phone'],
        [5.692, 'Standard Payment', 4.819, 5.373, 'coin'],
        [9.193, 'Debit/Credit Card', 8.322, 8.876, 'card']
    ];
    cols.forEach(([x, cap, bx, lx, kind]) => {
        sl.addShape('ellipse', { x, y: 3.359, w: 1.448, h: 1.465, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
        if (kind === 'phone') iconPhone(sl, x + 0.395, 3.68, 0.66);
        if (kind === 'coin') iconCoinHand(sl, x + 0.33, 3.72, 0.79);
        if (kind === 'card') iconCard(sl, x + 0.395, 3.72, 0.66);
        sl.addText(cap, base({ x: lx, y: 5.271, w: 2.082, h: 0.337, fontFace: F.med, fontSize: 14, color: C.ink2, align: 'center' }));
        body(sl, bx, 5.643, 3.19, 0.897, LOREM_TINY, { align: 'center' });
    });
});

/* 17 - split percentage card */
slide((sl) => {
    chrome(sl, 17, 3);
    sl.addShape('roundRect', { x: 1.0129, y: 1.327, w: 6.3557, h: 4.846, rectRadius: 0.2, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
    sl.addShape('line', { x: 4.1907, y: 1.327, w: 0, h: 4.846, line: { color: C.pale, width: 2.5 } });
    photoSlot(sl, 1.7657, 1.709, 1.6441, 1.567, { shape: 'ellipse' });
    photoSlot(sl, 4.9636, 1.709, 1.6441, 1.567, { shape: 'ellipse' });

    [['92%', 'Overlay Services', 1.942, 1.547, 1.202], ['85%', 'Mobile Wallets', 5.143, 4.748, 4.403]].forEach(([n, cap, nx, lx, bx]) => {
        stat(sl, nx, 3.683, 1.291, 0.774, n, 40);
        sl.addText(cap, base({ x: lx, y: 4.436, w: 2.082, h: 0.337, fontFace: F.med, fontSize: 14, color: C.ink2, align: 'center' }));
        body(sl, bx, 4.892, 2.772, 0.897, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the..', { align: 'center' });
    });

    kicker(sl, 8.011, 1.79, 'Our Services', { w: 1.7 });
    title(sl, 8.01, 2.155, 3.52, 1.986, 'What are the Different Method of Digital Payment ?');
    body(sl, 8.0, 4.257, 3.53, 1.453, LOREM_LONG.replace('it to make a type specimen book.', '') + ' ');
});

/* 18 - break slide */
slide((sl) => {
    gradRect(sl, 0, 1.4792, 4.3898, 6.0208, G.deep, 45);
    photoSlot(sl, 1.3777, 0, 4.8597, 6.5833);
    sl.addShape('roundRect', { x: 7.0951, y: 1.4754, w: 5.3547, h: 5.4379, rectRadius: 0.22, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
    gradRound(sl, 9.3958, 1.2841, 3.9375, 0.5657, 0.09, G.deep, 45, ['tl', 'bl']);

    sl.addText('Pinnacle', base({ x: -1.1111, y: 4.3895, w: 3.6345, h: 1.1107, fontFace: F.displaySemi, fontSize: 60, color: C.pale, rotate: 270, margin: 0 }));
    sl.addText('Let\u2019s Take A 5 \u2013 Minute Break', base({
        x: 9.6, y: 1.334, w: 3.53, h: 0.467, fontFace: F.med, fontSize: 16, color: C.pale,
        align: 'center', lineSpacingMultiple: 1.5, margin: 0
    }));
    sl.addText('Break', base({ x: 7.883, y: 2.56, w: 3.778, h: 1.582, fontFace: F.display, fontSize: 88, color: C.grad2, margin: 0 }));
    sl.addText('Slide', base({ x: 7.961, y: 3.872, w: 3.326, h: 1.582, fontFace: F.display, fontSize: 88, color: C.grad2, margin: 0 }));
    pill(sl, 10.556, 6.104, 1.427, 0.334, 'Learn More');
});

/* 19 - portfolio milestones */
slide((sl) => {
    chrome(sl, 19, 0);
    sl.addShape('roundRect', { x: 0.5414, y: 0.5572, w: 6.4709, h: 6.36, rectRadius: 0.19, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
    photoSlot(sl, 1.0056, 1.1997, 5.5424, 2.2632, { shape: 'roundRect', rectRadius: 0.24 });

    kicker(sl, 1.542, 3.881, 'Our Portfolio', { w: 1.6 });
    sl.addText([
        { text: 'Cost Savings', options: { bold: true } },
        { text: '\u00A0Through Greater Efficiency and Speed' }
    ], base({ x: 1.522, y: 4.245, w: 4.129, h: 1.313, fontFace: F.semi, fontSize: 24, color: C.ink }));
    body(sl, 1.544, 5.708, 4.108, 0.62, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the.');
    sl.addShape('ellipse', { x: 6.647, y: 3.954, w: 0.676, h: 0.691, fill: { color: C.yellow }, line: { type: 'none' } });
    sl.addText('\u203A', { x: 6.647, y: 3.9, w: 0.676, h: 0.691, fontFace: F.semi, fontSize: 26, color: C.white, align: 'center', valign: 'middle', margin: 0 });

    const rows = [
        ['1200+', 'Point of Sales', 1.017, C.grad1, C.ink2, C.body],
        ['700+', 'Payment Instrumen', 2.964, C.hair, C.mute, C.mute],
        ['2500+', 'Financial Inclusion', 4.993, C.hair, C.mute, C.mute]
    ];
    rows.forEach(([n, cap, y, nc, lc, bc]) => {
        sl.addText(n, base({ x: 7.96, y, w: 1.758, h: 0.639, fontFace: F.semi, fontSize: 32, color: nc }));
        sl.addText(cap, base({ x: 7.981, y: y + 0.607, w: 2.446, h: 0.337, fontFace: F.med, fontSize: 14, color: lc }));
        body(sl, 8.002, y + 0.952, 3.912, 0.62, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has. ', { color: bc });
    });
    [2.836, 4.757].forEach((y) => sl.addShape('line', { x: 8.07, y, w: 3.744, h: 0, line: { color: C.pale, width: 1 } }));
});

/* 20 - keys to a successful transaction */
slide((sl) => {
    chrome(sl, 20, 1);
    photoSlot(sl, 1.0, 1.0175, 3.3757, 4.7544);

    kicker(sl, 4.931, 1.178, 'Our Portfolio', { w: 1.6 });
    title(sl, 4.929, 1.543, 7.095, 0.976, 'PLACEHOLDER', { fontSize: 26 });
    body(sl, 4.945, 2.708, 3.241, 1.453, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer took a galley of type.");
    body(sl, 8.56, 2.708, 3.241, 1.453, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer took a galley of type.");

    sl.addText('Medium of Transmission', base({ x: 1.374, y: 6.114, w: 2.627, h: 0.337, fontFace: F.semi, fontSize: 14, color: C.ink2, align: 'center' }));
    sl.addText('Digital Payments Offer Significant Benefits', base({ x: 1.374, y: 6.418, w: 2.627, h: 0.471, fontFace: F.reg, fontSize: 11, color: C.body, align: 'center' }));

    const thumbs = [
        [5.0463, 1.9247, 'Fit-For-Purpose', 5.099, 1.82],
        [7.4375, 1.9325, 'Partially Digital', 7.427, 1.955],
        [9.8621, 1.9448, 'Digital Bank Transfer', 9.861, 1.955]
    ];
    thumbs.forEach(([x, w, cap, tx, tw]) => {
        photoSlot(sl, x, 4.5926, w, 1.628);
        gradRect(sl, x, 6.2051, w, 0.3791, G.card, 45, 8);
        sl.addText(cap, base({ x: tx, y: 6.252, w: tw, h: 0.286, fontFace: F.med, fontSize: 11, color: C.pale, align: 'center' }));
    });
});

/* 21 - gradient panel + numbered list */
slide((sl) => {
    chrome(sl, 21, 2);
    gradRound(sl, 0, 1.498, 6.282, 4.504, 0.3, G.panel, 45, ['tr', 'br']);

    kicker(sl, 0.906, 3.558, 'Our Portfolio', { w: 1.6, color: C.hair });
    title(sl, 0.896, 3.923, 3.913, 1.582, 'PLACEHOLDER', { fontSize: 22, color: C.white });

    [['01.', 'Cashless Economies', 1.21, 2.221], ['02.', 'Mobile Channels', 3.078, 1.967], ['03.', 'Online Banking Method', 4.939, 2.586]]
        .forEach(([num, head, y, hw]) => {
            sl.addText(num, base({ x: 7.159, y, w: 1.322, h: 0.909, fontFace: F.semi, fontSize: 48, color: C.grad1 }));
            sl.addText(head, base({ x: 8.382, y: y + 0.109, w: hw, h: 0.337, fontFace: F.med, fontSize: 14, color: C.ink2 }));
            body(sl, 8.384, y + 0.453, 3.53, 0.897, LOREM_SHORT);
        });
});

/* 22 - increasingly cashless */
slide((sl) => {
    chrome(sl, 22, 3);
    photoSlot(sl, 5.729, 1.1648, 6.083, 3.4317);

    kicker(sl, 0.902, 1.53, 'Our Portfolio', { w: 1.6 });
    title(sl, 0.883, 1.894, 3.911, 0.808, 'Payments Are Becoming Increasingly Cashless', { fontSize: 21 });
    body(sl, 0.898, 2.779, 3.911, 1.453, LOREM_LONG);

    [[1.012, 'Debit/Credit Card', 2.768, 3.349, 'card'], [6.667, 'Mobile Wallet', 8.423, 3.471, 'phone']]
        .forEach(([x, cap, tx, tw, kind]) => {
            sl.addShape('rect', { x, y: 5.218, w: 1.246, h: 1.276, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
            if (kind === 'card') iconCard(sl, x + 0.33, 5.56, 0.583);
            else iconPhone(sl, x + 0.33, 5.56, 0.583);
            sl.addText(cap, base({ x: tx, y: 5.234, w: 2.543, h: 0.337, fontFace: F.med, fontSize: 14, color: C.ink2 }));
            body(sl, tx + 0.02, 5.579, tw, 0.897, LOREM_SHORT);
        });
});

/* 23 - transaction value */
slide((sl) => {
    chrome(sl, 23, 0);
    gradRect(sl, 0.5477, -0.0096, 3.5379, 7.5192, G.steel, 45);
    photoSlot(sl, 8.8074, 1.2671, 3.0051, 4.9552);
    [0.7967, 4.0429].forEach((y) => photoSlot(sl, 1.4947, y, 1.6441, 1.567, { shape: 'ellipse' }));

    [['01.', 1.237, 1.259, 2.56], ['02.', 1.231, 4.458, 5.808]].forEach(([n, x, y, ty]) => {
        sl.addShape('ellipse', { x, y, w: 0.627, h: 0.642, fill: { color: C.yellow }, line: { type: 'none' } });
        sl.addText(n, { x, y, w: 0.627, h: 0.642, fontFace: F.med, fontSize: 12, color: C.pale, align: 'center', valign: 'middle', margin: 0 });
        body(sl, 1.022, ty, 2.588, 0.897, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. ', { align: 'center', color: C.hair });
    });

    kicker(sl, 4.648, 1.23, 'Our Portfolio', { w: 1.6 });
    title(sl, 4.63, 1.595, 3.603, 1.717, 'The Transaction Value for the Global Digital Payment Market', { fontSize: 24 });
    label(sl, 4.63, 3.581, 2.663, 'Research and Markets');
    body(sl, 4.644, 3.893, 3.603, 0.897, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever.");
    label(sl, 4.644, 5.06, 2.663, 'Driving Innovation');
    body(sl, 4.657, 5.373, 3.603, 0.897, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever.");
});

/* 24 - supporting digital economies */
slide((sl) => {
    chrome(sl, 24, 1);
    photoSlot(sl, 6.0862, 0, 5.7207, 6.1966);
    photoSlot(sl, 0, 5.1875, 5.0208, 2.3125);
    gradRect(sl, 5.018, 5.1875, 4.628, 2.3125, G.steel, 45);

    kicker(sl, 0.896, 1.0, 'Our Portfolio', { w: 1.6 });
    title(sl, 0.877, 1.325, 4.153, 1.851, 'Payments Are Supporting the Development of Digital Economies', { fontSize: 26 });
    body(sl, 0.898, 3.345, 4.229, 1.175,
        "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer");

    sl.addShape('ellipse', { x: 5.669, y: 5.516, w: 0.546, h: 0.552, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
    iconCoinHand(sl, 5.79, 5.64, 0.3);
    label(sl, 5.568, 6.176, 2.082, 'Standard Payment', { color: C.pale });
    body(sl, 5.566, 6.551, 3.532, 0.62, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has.', { align: 'left', color: C.hair });

    sl.addShape('ellipse', { x: 9.314, y: 5.998, w: 0.676, h: 0.691, fill: { color: C.yellow }, line: { type: 'none' } });
    sl.addText('\u203A', { x: 9.314, y: 5.944, w: 0.676, h: 0.691, fontFace: F.semi, fontSize: 26, color: C.white, align: 'center', valign: 'middle', margin: 0 });
});

/* 25..28 - SWOT family: gradient wedge on the left with a big translucent letter */
function swotSlide(sl, num, activeDot, letter, heading, extras) {
    chrome(sl, num, activeDot);
    // slanted gradient panel: full-height on the left, cut away to the right
    gradRect(sl, 0, 0, 5.533, 7.5, G.swot, 45);
    sl.addShape('custGeom', {
        x: 3.257, y: 0, w: 2.276, h: 7.5,
        points: [
            { x: 0, y: 0, moveTo: true }, { x: 2.276, y: 7.5 }, { x: 2.276, y: 0 }, { close: true }
        ],
        fill: { color: C.white }, line: { type: 'none' }
    });
    sl.addShape('ellipse', { x: 1.516, y: 1.528, w: 4.475, h: 4.433, fill: { color: 'BFBFBF', transparency: 11 }, line: { color: C.white, width: 4 } });
    sl.addText(letter, {
        x: 1.516, y: 1.44, w: 4.475, h: 4.433, fontFace: F.semi, fontSize: 210, bold: true,
        color: '7C57BE', align: 'center', valign: 'middle', margin: 0, outline: { color: C.white, size: 2 }
    });

    kicker(sl, 7.038, 1.697, 'SWOT Analysis', { w: 1.9 });
    title(sl, 7.03, 2.108, 4.581, 1.043, heading);
    extras(sl);
}

slide((sl) => swotSlide(sl, 25, 2, 'S', 'Strength Pinnacle Analysis Slides', (s) => {
    body(s, 7.042, 3.341, 4.586, 1.731, LOREM_LONG + '  It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged.');
    pill(s, 7.148, 5.406, 1.427, 0.334, 'Learn More');
}));

slide((sl) => swotSlide(sl, 26, 3, 'W', 'Weakness Pinnacle Analysis Slides', (s) => {
    body(s, 7.044, 3.341, 4.588, 1.453, LOREM_LONG + '  It has survived not only five centuries, but also the leap into electronic.');
    [['4500+', 'Credit / Debit Card', 7.141, 7.864, 7.849, 'card'], ['7300+', 'Mobile Wallets', 9.613, 10.329, 10.315, 'phone']]
        .forEach(([n, cap, ox, nx, lx, kind]) => {
            s.addShape('ellipse', { x: ox, y: 5.287, w: 0.57, h: 0.577, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
            if (kind === 'card') iconCard(s, ox + 0.155, 5.42, 0.26);
            else iconPhone(s, ox + 0.155, 5.42, 0.26);
            s.addText(n, base({ x: nx, y: 5.153, w: 1.2, h: 0.467, fontFace: F.semi, fontSize: 16, color: C.pink }));
            s.addText(cap, base({ x: lx, y: 5.488, w: 1.7, h: 0.375, fontFace: F.body, fontSize: 12, color: C.body }));
        });
}));

slide((sl) => swotSlide(sl, 27, 0, 'O', 'Opportunity Pinnacle Analysis Slides', (s) => {
    s.addShape('ellipse', { x: 7.15, y: 3.459, w: 0.898, h: 0.898, fill: { color: C.white }, line: { type: 'none' } });
    iconCoinHand(s, 7.362, 3.63, 0.48);
    s.addText('+5500', base({ x: 8.305, y: 3.488, w: 2.111, h: 0.842, fontFace: F.semi, fontSize: 44, color: C.grad2 }));
    body(s, 7.052, 4.701, 4.559, 1.175,
        "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s. when an unknown printer took a galley of type and scrambled it to make a type specimen book. ");
}));

slide((sl) => swotSlide(sl, 28, 1, 'T', 'Threats Pinnacle Analysis Slides', (s) => {
    [['01', 'Fully Automated', 3.311, 3.497, 3.627], ['02', 'Cashless Global Society', 4.653, 4.845, 4.975]]
        .forEach(([n, cap, ty, oy, by]) => {
            s.addShape('ellipse', { x: 7.147, y: oy, w: 0.54, h: 0.542, fill: { color: C.yellow }, line: { type: 'none' }, shadow: SHADOW });
            s.addText(n, { x: 7.147, y: oy, w: 0.54, h: 0.542, fontFace: F.med, fontSize: 10.5, color: C.white, align: 'center', valign: 'middle', margin: 0 });
            s.addText('\u00A0' + cap, base({ x: 7.889, y: ty, w: 2.2, h: 0.375, fontFace: F.med, fontSize: 12, color: C.ink2 }));
            body(s, 7.929, by, 3.682, 0.897, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever.");
        });
}));

/* 29 - laptop mockup */
slide((sl) => {
    chrome(sl, 29, 2);
    gradRect(sl, 0, 0, 3.1538, 7.5, G.card, 45);
    imagePlaceholder(sl, 0, 1.2895, 6.4865, 4.9698, { color: 'DCD9E2' });

    kicker(sl, 7.038, 1.395, 'Mockup Devices', { w: 2.1 });
    title(sl, 7.029, 1.772, 4.847, 1.515, 'Where Are We in the Digital Payment Space?');
    body(sl, 7.059, 3.439, 4.847, 1.175, LOREM_LONG);

    gradRound(sl, 7.1529, 4.913, 2.1035, 1.1915, 0.08, G.band, 45);
    sl.addText('42681', base({ x: 7.604, y: 5.065, w: 1.228, h: 0.538, fontFace: F.semi, fontSize: 26, color: C.pale, align: 'center', margin: 0 }));
    sl.addText('\u00A0Sales Terminals', base({ x: 7.043, y: 5.616, w: 2.322, h: 0.337, fontFace: F.light, fontSize: 14, color: C.hair, align: 'center' }));
    sl.addShape('ellipse', { x: 8.919, y: 5.011, w: 0.243, h: 0.243, fill: { color: C.yellow }, line: { type: 'none' } });
    iconCheck(sl, 9.04, 5.132, 0.19);

    sl.addShape('roundRect', { x: 9.6726, y: 4.913, w: 2.1035, h: 1.1915, rectRadius: 0.08, fill: { color: C.pale }, line: { color: C.hair, width: 1.5 } });
    sl.addText('8730', base({ x: 10.202, y: 5.063, w: 1.068, h: 0.538, fontFace: F.semi, fontSize: 26, color: C.mute, align: 'center', margin: 0 }));
    sl.addText('Great Upheaval', base({ x: 9.568, y: 5.615, w: 2.322, h: 0.337, fontFace: F.light, fontSize: 14, color: C.mute, align: 'center' }));
    sl.addShape('line', { x: 9.9959, y: 6.077, w: 1.4569, h: 0, line: { color: C.hair, width: 3 } });
});

/* 30 - tablet mockup */
slide((sl) => {
    chrome(sl, 30, 3);
    imagePlaceholder(sl, 4.1998, 0.8368, 3.672, 5.8263, { color: 'D2CFD8' });
    [0.9852, 4.0286].forEach((y) => photoSlot(sl, 1.5213, y, 1.3869, 1.3389, { shape: 'ellipse' }));

    [['\u00A0Multiple Ecosystems.', 2.526, 2.866], ['Payment Landscape', 5.717, 6.057]].forEach(([cap, ly, by]) => {
        sl.addText(cap, base({ x: 0.859, y: ly, w: 2.712, h: 0.337, fontFace: F.med, fontSize: 14, color: C.ink2, align: 'center' }));
        body(sl, 0.783, by, 2.863, 0.62, 'Lorem Ipsum is simply dummy text of the printing and typesetting.', { align: 'center' });
    });

    kicker(sl, 8.59, 1.253, 'Mockup Devices', { w: 2.1 });
    title(sl, 8.586, 1.63, 3.272, 1.986, 'Making the Payment Either Entirely in the Background');

    gradRound(sl, 8.7088, 3.9451, 3.1101, 1.0816, 0.13, G.band, 45);
    sl.addShape('ellipse', { x: 8.867, y: 4.197, w: 0.571, h: 0.578, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
    iconBars(sl, 9.0, 4.33, 0.31);
    label(sl, 9.514, 4.159, 2.234, 'Overlay Services', { color: C.pale });
    body(sl, 9.52, 4.471, 2.234, 0.342, 'Lorem Ipsum is simply dummy', { color: C.hair });

    sl.addShape('ellipse', { x: 8.869, y: 5.417, w: 0.571, h: 0.578, fill: { color: C.pale }, line: { type: 'none' }, shadow: SHADOW });
    iconCoinHand(sl, 8.975, 5.53, 0.36, C.mute);
    label(sl, 9.516, 5.38, 2.234, 'Standard Payment', { color: C.mute });
    body(sl, 9.521, 5.691, 2.234, 0.342, 'Lorem Ipsum is simply dummy', { color: C.mute });
});

/* 31 - phone mockup collage */
slide((sl) => {
    chrome(sl, 31, 0);
    gradRect(sl, -0.014, 0, 6.979, 7.5, G.tile, 45);
    // tiled phone mockups behind the hero device
    [[0.1319, -0.9223, 2.6942, 4.063], [5.5718, 0.7009, 2.6942, 4.9309], [3.5664, 5.2827, 2.6942, 3.0789],
     [0.1903, 5.6458, 2.6942, 2.5454], [-0.9379, 2.9954, 1.8397, 3.2672], [3.8954, -0.8038, 2.6942, 2.2207]]
        .forEach(([x, y, w, h]) => sl.addShape('roundRect', {
            x, y, w, h, rotate: 29, rectRadius: 0.1,
            fill: { color: C.white, transparency: 92 }, line: { type: 'none' }
        }));
    imagePlaceholder(sl, 2.107, 1.0745, 2.6942, 4.9827, { rotate: 29, color: 'CFCBD8' });

    kicker(sl, 7.566, 1.249, 'Mockup Devices', { w: 1.95, fontSize: 14, h: 0.34 });
    title(sl, 7.564, 1.592, 4.068, 1.043, 'The New Mobile Payment Solution');
    body(sl, 7.564, 2.715, 4.361, 0.97, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever.", { fontSize: 12 });

    [['Standard Payment', 4.036, 4.332, 4.139, 'coin'], ['Credit / Debit Card ', 5.288, 5.584, 5.459, 'card']]
        .forEach(([cap, ly, by, oy, kind]) => {
            sl.addShape('ellipse', { x: 7.665, y: oy, w: 0.637, h: 0.637, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
            if (kind === 'coin') iconCoinHand(sl, 7.81, oy + 0.15, 0.35);
            else iconCard(sl, 7.81, oy + 0.16, 0.35);
            label(sl, 8.61, ly, 2.117, cap);
            body(sl, 8.61, by, 3.32, 0.667, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.', { fontSize: 12, align: 'left' });
        });
});

/* 32 - pinwheel infographic */
slide((sl) => {
    chrome(sl, 32, 1);
    kicker(sl, 5.458, 0.844, 'Mockup Devices', { w: 1.949, fontSize: 14, align: 'center', h: 0.34 });
    title(sl, 3.763, 1.23, 5.338, 0.555, 'Infographic Process Slide', { fontSize: 27, align: 'center', charSpacing: 0.5 });

    /*
     * Pinwheel: four blades spaced 90 degrees around the hub.  A blade is a
     * violet half-disc plus a gold quarter-disc sharing the same centre, the
     * gold sitting in the quadrant just clockwise of the half-disc's flat edge.
     */
    const hub = [6.4165, 4.267];
    const ARM = 1.066;   // hub -> blade centre
    const R = 0.867;     // blade radius
    [[270, '01'], [180, '02'], [0, '03'], [90, '04']].forEach(([deg, num]) => {
        const rad = deg * Math.PI / 180;
        const px = hub[0] + ARM * Math.cos(rad);
        const py = hub[1] + ARM * Math.sin(rad);
        ring(sl, px, py, R, R, 180 + deg, 180, (t) => stopColor(G.card, t), 10);
        ring(sl, px, py, R * 0.94, R * 0.94, deg, 90, C.yellow, 1);
        const q = (deg + 45) * Math.PI / 180;
        sl.addText(num, {
            x: px + R * 0.5 * Math.cos(q) - 0.25, y: py + R * 0.5 * Math.sin(q) - 0.2, w: 0.5, h: 0.4,
            fontFace: F.semi, fontSize: 18, color: C.white, align: 'center', valign: 'middle', margin: 0
        });
    });

    const items = [
        ['01. Debit / Credit ', 0.678, 2.85, 0.882, 3.294, 'right'],
        ['02. Overlay Services', 0.668, 4.927, 0.882, 5.318, 'right'],
        ['03. Standard Payment', 9.1, 2.846, 9.002, 3.291, 'left'],
        ['04. Mobile Wallets', 9.1, 4.924, 9.002, 5.315, 'left']
    ];
    items.forEach(([t, hx, hy, bx, by, al]) => {
        sl.addText(t, base({ x: hx, y: hy, w: 3.12, h: 0.37, fontFace: F.med, fontSize: 16, color: C.ink2, align: al, charSpacing: 0.5, valign: 'bottom' }));
        body(sl, bx, by, 2.99, 0.97, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been.', { fontSize: 12 });
    });
});

/* 33 - four-ring process */
slide((sl) => {
    chrome(sl, 33, 2);
    kicker(sl, 5.458, 0.844, 'Mockup Devices', { w: 1.949, fontSize: 14, align: 'center', h: 0.34 });
    title(sl, 3.763, 1.23, 5.338, 0.555, 'Infographic Process Slide', { fontSize: 27, align: 'center', charSpacing: 0.5 });

    const rings = [
        ['Overlay', 2.164, 'bars', 0.884, 0.991],
        ['Mobile', 4.996, 'phone', 3.754, 3.861],
        ['Credit/Debit', 7.827, 'card', 6.588, 6.695],
        ['Payment', 10.659, 'coin', 9.423, 9.53]
    ];
    const cy = 3.479;
    rings.forEach(([cap, cx, kind, lx, bx], i) => {
        // connector between neighbouring rings
        if (i < 3) gradRect(sl, cx + 0.677, cy - 0.066, 0.678, 0.133, i % 2 ? [[0, C.blue], [1, C.pink]] : [[0, C.pink], [1, C.blue]], 0, 8);
        // gold half of the ring then the gradient half on top
        ring(sl, cx, cy, 1.128, 0.42, i % 2 ? -180 : 0, 180, C.yellow);
        ring(sl, cx, cy, 1.128, 0.42, i % 2 ? 0 : 180, 180, (t) => stopColor(G.card, i % 2 ? 1 - t : t), 10);
        if (kind === 'bars') iconBars(sl, cx - 0.42, cy - 0.36, 0.78);
        if (kind === 'phone') iconPhone(sl, cx - 0.39, cy - 0.39, 0.78);
        if (kind === 'card') iconCard(sl, cx - 0.42, cy - 0.34, 0.84);
        if (kind === 'coin') iconCoinHand(sl, cx - 0.42, cy - 0.4, 0.84);
        sl.addText(cap, base({ x: lx, y: 4.904, w: 2.478, h: 0.438, fontFace: F.med, fontSize: 20, color: C.ink2, align: 'center', charSpacing: 0.5, valign: 'bottom' }));
        body(sl, bx, 5.37, 2.28, 0.897, 'Lorem Ipsum sily dummy \ntext of the printing and typesetting industry.', { align: 'center' });
    });
});

/* 34 - zig-zag bubbles */
slide((sl) => {
    chrome(sl, 34, 3);
    kicker(sl, 5.458, 0.844, 'Mockup Devices', { w: 1.949, fontSize: 14, align: 'center', h: 0.34 });
    title(sl, 3.763, 1.23, 5.338, 0.555, 'Infographic Process Slide', { fontSize: 27, align: 'center', charSpacing: 0.5 });

    // W-shaped ribbon threaded through the bubble centres
    const nodes = [[3.075, 3.214], [4.866, 4.664], [6.44, 3.932], [8.041, 4.679], [9.812, 3.213]];
    for (let i = 0; i < nodes.length - 1; i++) {
        const [x1, y1] = nodes[i];
        const [x2, y2] = nodes[i + 1];
        const len = Math.hypot(x2 - x1, y2 - y1);
        sl.addShape('rect', {
            x: (x1 + x2) / 2 - len / 2, y: (y1 + y2) / 2 - 0.26, w: len, h: 0.52,
            rotate: Math.atan2(y2 - y1, x2 - x1) * 180 / Math.PI,
            fill: { color: i < 2 ? 'A94AA0' : '9A48A3' }, line: { type: 'none' }
        });
    }

    const D = 1.717;
    [[3.075, 3.214, 'bars'], [4.866, 4.664, 'phone'], [8.041, 4.679, 'coin'], [9.812, 3.213, 'card']]
        .forEach(([cx, cy, kind]) => {
            gradRound(sl, cx - D / 2, cy - D / 2, D, D, D / 2, G.card, 45);
            if (kind === 'bars') iconBars(sl, cx - 0.36, cy - 0.36, 0.72, C.white);
            if (kind === 'phone') iconPhone(sl, cx - 0.36, cy - 0.36, 0.72, C.white);
            if (kind === 'coin') iconCoinHand(sl, cx - 0.4, cy - 0.38, 0.8, C.white);
            if (kind === 'card') iconCard(sl, cx - 0.4, cy - 0.32, 0.8, C.white);
        });
    // gold hub with the coin logo
    sl.addShape('ellipse', { x: 5.582, y: 3.073, w: 1.717, h: 1.717, fill: { color: C.yellow }, line: { type: 'none' } });
    coin(sl, 5.855, 3.345, 1.171);

    sl.addText('Most Importantly', base({ x: 0.814, y: 4.777, w: 2.202, h: 0.37, fontFace: F.med, fontSize: 16, color: C.ink2, align: 'right', valign: 'bottom' }));
    body(sl, 0.987, 5.145, 2.029, 1.175, LOREM_TINY);
    sl.addText('Connect Globally', base({ x: 9.77, y: 4.777, w: 2.219, h: 0.37, fontFace: F.med, fontSize: 16, color: C.ink2, valign: 'bottom' }));
    body(sl, 9.781, 5.145, 2.029, 1.175, LOREM_TINY);
});

/* 35 - pie chart */
slide((sl) => {
    chrome(sl, 35, 0);
    sl.addShape('roundRect', { x: 0.7451, y: 2.3509, w: 4.8823, h: 4.3528, rectRadius: 0.14, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });

    kicker(sl, 5.692, 0.844, 'Our Chart', { w: 1.949, fontSize: 14, align: 'center', h: 0.34 });
    title(sl, 4.169, 1.23, 4.996, 1.043, 'Design Would Be Tied to A Digital Identity', { align: 'center' });
    body(sl, 6.577, 2.73, 5.287, 0.62, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever .");

    sl.addChart('pie', [{
        name: 'Sales',
        labels: ['Intregity', 'Finality', 'Liqudity', 'Security'],
        values: [8.2, 3.2, 1.4, 1.2]
    }], {
        x: 1.034, y: 2.726, w: 4.304, h: 3.603,
        chartColors: [C.pink, C.violet, C.blue, C.yellow],
        dataBorder: { pt: 0.75, color: C.white },
        showLegend: true, legendPos: 'b', legendFontFace: F.med, legendFontSize: 11, legendColor: C.body,
        showValue: false, showTitle: false, firstSliceAng: 0
    });

    [['01.   Intregity', 6.56, 3.731, 6.926, 4.044], ['02.  Finality', 6.56, 5.195, 6.926, 5.508],
     ['03.  Liqudity', 9.479, 3.731, 9.861, 4.05], ['04.  Security', 9.48, 5.195, 9.861, 5.514]]
        .forEach(([t, lx, ly, bx, by]) => {
            label(sl, lx, ly, 2.663, t);
            body(sl, bx, by, 2.17, 0.62, 'Lorem Ipsum is simply dummy text of the printing.', { align: 'left' });
        });
});

/* 36 - bar chart */
slide((sl) => {
    chrome(sl, 36, 1);
    sl.addShape('roundRect', { x: 6.8958, y: 0.7884, w: 5.1042, h: 4.1754, rectRadius: 0.135, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
    sl.addShape('roundRect', { x: 5.4375, y: 4.4178, w: 2.1974, h: 2.2939, rectRadius: 0.135, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });

    kicker(sl, 0.896, 1.776, 'Our Charts', { w: 1.5 });
    title(sl, 0.877, 2.101, 3.639, 1.986, 'Working to Safeguard Public Trust In Money and Payments');
    label(sl, 0.892, 4.276, 2.657, 'Infrastructure Providers');
    body(sl, 0.898, 4.549, 3.618, 1.175,
        "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer");

    sl.addChart('bar', [
        { name: 'Series 1', labels: ['Liqudity', 'Finality', 'Security', 'Intregity'], values: [4.3, 2.5, 3.5, 4.5] },
        { name: 'Series 2', labels: ['Liqudity', 'Finality', 'Security', 'Intregity'], values: [2.4, 4.4, 1.8, 2.8] }
    ], {
        x: 7.247, y: 1.201, w: 4.402, h: 3.35,
        barDir: 'col', barGapWidthPct: 219, barOverlapPct: -27,
        chartColors: [C.violet, C.pink],
        showLegend: false, showValue: false, showTitle: false,
        catAxisLineShow: true, catAxisLineColor: 'D9D9D9',
        catAxisLabelFontFace: F.med, catAxisLabelFontSize: 11, catAxisLabelColor: C.body,
        valAxisHidden: true, valGridLine: { style: 'solid', color: 'D9D9D9', size: 0.75 }
    });

    sl.addShape('ellipse', { x: 6.218, y: 4.765, w: 0.637, h: 0.637, fill: { color: '7B41AB' }, line: { type: 'none' } });
    iconCheck(sl, 6.532, 5.082, 0.35, C.pale);
    sl.addText('Digital Assets', base({ x: 5.614, y: 5.566, w: 1.843, h: 0.286, fontFace: F.med, fontSize: 11, color: C.ink, align: 'center' }));
    body(sl, 5.721, 5.836, 1.63, 0.62, 'Lorem Ipsum  simply\n dummy text  printing', { align: 'center' });
    body(sl, 8.179, 5.615, 3.321, 0.62,
        '\u201CLorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has.\u201D', { italic: true });
});

/* 37 - pricing table */
slide((sl) => {
    chrome(sl, 37, 2);
    kicker(sl, 5.458, 0.844, 'Our Pricing Table', { w: 1.949, fontSize: 14, align: 'center', h: 0.34 });
    title(sl, 3.299, 1.23, 6.271, 0.555, 'Product Pricing and Licensing', { fontSize: 27, align: 'center', charSpacing: 0.5 });

    const plans = [
        ['Basic Plan', '$50 ', 1.0027, 1.292, 1.282, 1.308, 1.675, 1.396, 'shield', 2],
        ['Standart Plan', '$70 ', 4.8773, 5.181, 5.171, 5.197, 5.549, 5.285, 'layers', 3],
        ['Advanced Plan', '$85 ', 8.7951, 9.095, 9.085, 9.111, 9.467, 9.199, 'card', 4]
    ];
    const feats = ['Set Instantly', 'Type of Product', 'International Growth', 'Credit Card'];
    plans.forEach(([name, price, cx, nx, px, fx, bx, ox, icon, enabled]) => {
        sl.addShape('roundRect', { x: cx, y: 2.2205, w: 3.0121, h: 4.4462, rectRadius: 0.146, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
        sl.addShape('ellipse', { x: ox, y: 2.704, w: 0.664, h: 0.658, fill: { color: C.white }, line: { type: 'none' }, shadow: SHADOW });
        if (icon === 'shield') sl.addShape('flowChartExtract', { x: ox + 0.185, y: 2.858, w: 0.294, h: 0.35, rotate: 180, fill: { color: C.violet }, line: { type: 'none' } });
        if (icon === 'layers') iconBars(sl, ox + 0.13, 2.87, 0.4);
        if (icon === 'card') iconCard(sl, ox + 0.145, 2.85, 0.376);

        sl.addText(name, base({ x: nx, y: 3.506, w: 2.05, h: 0.37, fontFace: F.med, fontSize: 16, color: C.ink2 }));
        sl.addText([
            { text: price, options: { fontSize: 28, color: C.pink } },
            { text: '/Month', options: { fontSize: 18, color: C.ink2 } }
        ], base({ x: px, y: 3.995, w: 2.074, h: 0.572, fontFace: F.med }));
        feats.forEach((f, i) => {
            sl.addText(f, base({
                x: fx, y: 4.754 + i * 0.3755, w: 2.43, h: 0.303, fontFace: F.med, fontSize: 12,
                charSpacing: 1, color: i < enabled ? C.ink2 : C.mute, bullet: { characterCode: '2022' }, margin: 0
            }));
        });
        pill(sl, bx, 6.471, 1.668, 0.39, 'GET STARTED', { font: F.med });
    });
});

/* 38 - contact */
slide((sl) => {
    chrome(sl, 38, 3);
    photoSlot(sl, 0.981, 0, 5.282, 6.1966);

    kicker(sl, 7.174, 1.143, 'Contact Us', { w: 1.949, fontSize: 14, h: 0.34 });
    title(sl, 7.174, 1.479, 4.752, 1.043, 'Let\u2019s Keep in Touch with Us? ');
    body(sl, 7.174, 2.665, 4.752, 0.62, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard.");

    sl.addShape('ellipse', { x: 7.247, y: 3.578, w: 0.549, h: 0.549, fill: { color: C.white }, line: { type: 'none' } });
    iconPin(sl, 7.435, 3.7, 0.22);
    sl.addText('Our Address', base({ x: 8.262, y: 3.498, w: 2.0, h: 0.353, fontFace: F.med, fontSize: 15, charSpacing: 0.5, color: C.ink }));
    body(sl, 8.26, 3.92, 3.677, 0.667, '1234 Budiono Street , IN 012  Soerakarta, Indonesia', { align: 'left', fontSize: 12 });

    sl.addShape('ellipse', { x: 7.247, y: 4.932, w: 0.549, h: 0.549, fill: { color: C.white }, line: { type: 'none' } });
    iconReceiver(sl, 7.406, 5.09, 0.232, C.pink);
    sl.addText('Phone & Email', base({ x: 8.266, y: 4.884, w: 2.2, h: 0.353, fontFace: F.med, fontSize: 15, charSpacing: 0.5, color: C.ink }));
    body(sl, 8.265, 5.306, 2.239, 0.667, '(0123) 4567 - 8910\nYouremail@payment.com', { align: 'left', fontSize: 12 });
});

/* 39 - quote */
slide((sl) => {
    photoSlot(sl, 0, 0, 13.3333, 7.5);
    gradRect(sl, 1.979, 1.429, 9.376, 4.641, G.quote, 45);
    sl.addText('\u201C', base({ x: 2.889, y: 1.611, w: 1.293, h: 2.895, fontFace: F.semi, fontSize: 166, color: C.yellow, margin: 0 }));
    sl.addText('These heroes of finance are like beads on a string, when one slips off, all the rest follow.', base({
        x: 4.058, y: 2.775, w: 6.233, h: 2.128, fontFace: F.quote, fontSize: 28, italic: true,
        color: C.pale, align: 'justify', lineSpacingMultiple: 1.5
    }));
    sl.addShape('line', { x: 3.188, y: 5.479, w: 4.75, h: 0, line: { color: C.pale, width: 1.5 } });
    sl.addText('Henrik Ibsen', base({ x: 8.544, y: 5.26, w: 1.746, h: 0.404, fontFace: F.med, fontSize: 18, color: C.white, align: 'right' }));
});

/* 40 - thank you */
slide((sl) => {
    sl.addShape('rect', { x: 1.3684, y: 0.54, w: 11.4458, h: 6.3857, fill: { type: 'none' }, line: { color: C.pale, width: 2 } });
    gradRect(sl, 0, 0, 4.1754, 7.5, G.deep, 45);
    photoSlot(sl, 1.9339, 1.5409, 4.4611, 4.4182, { shape: 'ellipse' });

    [['MARKETING', 0.338, 0.86, 1.133], ['FINANCE', 0.451, 3.5, 0.898], ['WALLET', 0.498, 6.498, 0.812]]
        .forEach(([t, x, y, w]) => sl.addText(t, base({
            x, y, w, h: 0.303, fontFace: F.med, fontSize: 12, color: C.white, align: 'center', rotate: 270, margin: 0
        })));
    [2.077, 2.324, 2.575, 4.861, 5.107, 5.358].forEach((y, i) => sl.addShape('ellipse', {
        x: i < 3 ? 0.84 : 0.839, y, w: 0.124, h: 0.124, fill: { type: 'none' }, line: { color: C.hair, width: 1.5 }
    }));

    sl.addText('Thank', base({ x: 7.611, y: 1.852, w: 3.962, h: 1.582, fontFace: F.display, fontSize: 88, color: C.grad2, align: 'center', margin: 0 }));
    sl.addText('You', base({ x: 8.308, y: 3.165, w: 2.569, h: 1.582, fontFace: F.display, fontSize: 88, color: C.grad2, align: 'center', margin: 0 }));
    sl.addShape('line', { x: 7.217, y: 4.915, w: 4.75, h: 0, line: { color: C.hair, width: 2.25 } });
    sl.addText('For Watching This Presentation', base({
        x: 7.719, y: 5.245, w: 3.746, h: 0.467, fontFace: F.med, fontSize: 16, color: C.ink2,
        align: 'center', lineSpacingMultiple: 1.5
    }));
});

/* --------------------------------------------------------------------- write */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'PINNACLE', width: 13.3333, height: 7.5 });
pptx.layout = 'PINNACLE';
pptx.title = 'Pinnacle - Digital Payment Presentation';

builders.forEach((build) => {
    const sl = pptx.addSlide();
    sl.background = { color: C.white };
    build(sl);
});

pptx.writeFile({ fileName: path.join(__dirname, '122f2088-c2ca-452e-9592-2fa816608a8a_grok_final.pptx') })
    .then((f) => console.log('wrote ' + f));
