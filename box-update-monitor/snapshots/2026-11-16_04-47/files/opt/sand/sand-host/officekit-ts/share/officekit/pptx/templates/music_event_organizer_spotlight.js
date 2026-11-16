/*
 * "SPOTLIGHT" music-event deck - rebuilt with pptxgenjs.
 * 30 slides, 13.333 x 7.5 in (16:9).
 *
 * Raster images of the original deck (device mock-ups / photos) are replaced by
 * flat placeholder rectangles labelled "[image]".  Gradients, which pptxgenjs
 * cannot express natively, are painted as a strip of interpolated solid shapes
 * (see gradRect / radialGlow).
 */

const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const BG = '171529'; // deck background (slide master)
const PANEL = '232134'; // slightly lighter card panel
const MAG = 'A700C2'; // magenta accent
const VIO = '5100B5'; // violet accent
const DEEP = '621BBA'; // deep purple (text on white chips)
const INK = '091211'; // near-black used on the cover slides
const W = 'FFFFFF';
const W95 = 'F2F2F2';
const W85 = 'D9D9D9';
const W75 = 'BFBFBF';
const GREY50 = '808080';
const GREY25 = '404040';

/* -------------------------------------------------------------------- fonts */
const PSB = 'Poppins SemiBold';
const PM = 'Poppins Medium';
const PL = 'Poppins Light';
const PR = 'Poppins';
const LATO = 'Lato';
const LATOB = 'Lato Black';
const MSB = 'Montserrat SemiBold';
const MBL = 'Montserrat Black';

/* ------------------------------------------------------------------ helpers */
const hex = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0').toUpperCase();

/** linear blend of two "RRGGBB" strings, t = 0 -> a, t = 1 -> b */
function mix(a, b, t) {
    let out = '';
    for (let i = 0; i < 3; i++) {
        const ca = parseInt(a.substr(i * 2, 2), 16);
        const cb = parseInt(b.substr(i * 2, 2), 16);
        out += hex(ca + (cb - ca) * t);
    }
    return out;
}

/**
 * Linear gradient faked with a fan of solid bands.
 *   dir 'h'  left  -> right      dir 'v'  top -> bottom
 *   dir 'd'  top-left -> bottom-right    dir 'a'  bottom-left -> top-right
 * The diagonals are a 50 % translucent vertical pass over a horizontal pass,
 * which is exactly a bilinear (diagonal) blend of the two colours.
 */
function gradRect(s, o) {
    const steps = o.steps || 24;
    const band = function (dir, c1, c2, alpha) {
        for (let i = 0; i < steps; i++) {
            const t = i / (steps - 1);
            const fill = { color: mix(c1, c2, t), transparency: alpha };
            if (dir === 'v') {
                s.addShape('rect', { x: o.x, y: o.y + (i * o.h) / steps, w: o.w, h: o.h / steps + 0.012, fill: fill, line: { type: 'none' } });
            } else {
                s.addShape('rect', { x: o.x + (i * o.w) / steps, y: o.y, w: o.w / steps + 0.012, h: o.h, fill: fill, line: { type: 'none' } });
            }
        }
    };
    const dir = o.dir || 'h';
    const alpha = o.transparency || 0;
    if (dir === 'h' || dir === 'v') band(dir, o.c1, o.c2, alpha);
    else if (dir === 'd') { band('h', o.c1, o.c2, alpha); band('v', o.c1, o.c2, 50); }
    else { band('h', o.c1, o.c2, alpha); band('v', o.c2, o.c1, 50); }
}

/** Vertical gradient inside a circle/ellipse: bands clipped to the ellipse chord. */
function gradEllipse(s, o) {
    const steps = o.steps || 22;
    const rx = o.w / 2, ry = o.h / 2;
    for (let i = 0; i < steps; i++) {
        const t = i / (steps - 1);
        const dy = (i + 0.5) / steps * 2 - 1; // -1 .. 1 across the ellipse
        const half = rx * Math.sqrt(Math.max(0, 1 - dy * dy));
        s.addShape('rect', {
            x: o.x + rx - half, y: o.y + (i * o.h) / steps, w: half * 2, h: o.h / steps + 0.01,
            fill: { color: mix(o.c1, o.c2, t) }, line: { type: 'none' },
        });
    }
}

/** Bounding box of a rect that PowerPoint stores unrotated but draws at 90/270 deg. */
function rot90(x, y, w, h) {
    return { x: x + w / 2 - h / 2, y: y + h / 2 - w / 2, w: h, h: w };
}

/**
 * The deck's signature magenta -> violet panel.
 * dir 'a' (default) puts magenta bottom-left, violet top-right; 'd' is the mirror.
 */
function accentPanel(s, x, y, w, h, dir) {
    gradRect(s, { x: x, y: y, w: w, h: h, c1: MAG, c2: VIO, dir: dir || 'a' });
}

/** Radial "corner glow" faked with concentric opaque ellipses (paint before content). */
function radialGlow(s, o) {
    const steps = o.steps || 16;
    for (let i = 0; i < steps; i++) {
        const t = i / steps; // 0 = outermost ring
        const r = o.r * (1 - t);
        s.addShape('ellipse', {
            x: o.cx - r, y: o.cy - r, w: r * 2, h: r * 2,
            fill: { color: mix(o.edge || BG, o.core, Math.pow(t, 1.6)) },
            line: { type: 'none' },
        });
    }
}

/** Text with PowerPoint-ish defaults (top anchored, no bullet). */
function tx(s, text, o) {
    s.addText(text, Object.assign({ valign: 'top', color: W, fontFace: PR, fontSize: 12, isTextBox: true }, o));
}

/** Plain filled rectangle. */
function box(s, o) {
    s.addShape(o.shape || 'rect', {
        x: o.x, y: o.y, w: o.w, h: o.h,
        fill: o.fill === null ? { type: 'none' } : { color: o.fill, transparency: o.transparency || 0 },
        line: o.line || { type: 'none' },
        rotate: o.rotate, rectRadius: o.rectRadius, flipH: o.flipH, flipV: o.flipV,
    });
}

/** Stand-in for a raster image from the source deck. */
function imgBox(s, x, y, w, h, label) {
    box(s, { x: x, y: y, w: w, h: h, fill: '2A2740', line: { color: '3E3A5A', width: 1 } });
    tx(s, label || '[image]', {
        x: x, y: y + h / 2 - 0.16, w: w, h: 0.32,
        align: 'center', fontFace: PL, fontSize: 10, color: '6E6A8C',
    });
}

/* --------------------------------------------------------- small components */

/** Gradient "Read More" pill (1.02 x 0.30). */
function readMore(s, x, y) {
    box(s, { x: x, y: y, w: 1.016, h: 0.303, fill: MAG, rectRadius: 0.05 });
    box(s, { x: x + 0.42, y: y, w: 0.596, h: 0.303, fill: VIO, rectRadius: 0.05 });
    box(s, { x: x + 0.42, y: y, w: 0.16, h: 0.303, fill: mix(MAG, VIO, 0.5) });
    tx(s, 'Read More', { x: x, y: y + 0.03, w: 1.016, h: 0.24, align: 'center', fontFace: PM, fontSize: 8, color: W95 });
}

/** Rounded square carousel button holding a chevron. */
function navBtn(s, o) {
    box(s, { x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill || W95, rectRadius: 0.08 });
    tx(s, o.ch || '>', {
        x: o.x, y: o.y, w: o.w, h: o.h, align: 'center', valign: 'middle',
        fontFace: PM, fontSize: o.size || 20, bold: true, color: o.color || MAG,
    });
}

/** Dark stat card with a big number (slides 4, 20, 25). */
function statCard(s, o) {
    box(s, { x: o.x, y: o.y, w: o.w, h: o.h, fill: PANEL });
    tx(s, o.value, {
        x: o.x, y: o.y + (o.vy || 0.36), w: o.w, h: 0.62, align: 'center',
        fontFace: PSB, fontSize: o.vsize || 30, color: o.vcolor || W95,
    });
    if (o.label) {
        tx(s, o.label, {
            x: o.x, y: o.y + (o.ly || 0.98), w: o.w, h: 0.28, align: 'center',
            fontFace: PM, fontSize: 10, color: W95,
        });
    }
}

/** Segmented circular progress gauge (slide 5). */
function gauge(s, o) {
    const N = 18, lit = Math.round((N * o.pct) / 100);
    for (let i = 0; i < N; i++) {
        const ang = (i * 360) / N;
        const rad = ((ang - 90) * Math.PI) / 180;
        const seg = 0.30 * o.r;
        const cx = o.cx + Math.cos(rad) * o.r * 0.80;
        const cy = o.cy + Math.sin(rad) * o.r * 0.80;
        box(s, {
            x: cx - seg * 0.22, y: cy - seg / 2, w: seg * 0.44, h: seg,
            fill: i < lit ? VIO : 'C9A8DD', rotate: ang,
        });
    }
    box(s, { x: o.cx - o.r * 0.55, y: o.cy - o.r * 0.55, w: o.r * 1.1, h: o.r * 1.1, shape: 'ellipse', fill: '8E00BE' });
    tx(s, o.pct + '%', {
        x: o.cx - o.r, y: o.cy - 0.17, w: o.r * 2, h: 0.34,
        align: 'center', valign: 'middle', fontFace: PSB, fontSize: 20, bold: true, color: W,
    });
}

/* ------------------------------------------------------------------- icons  */
/* Simple, native-shape stand-ins for the deck's vector music icons.           */

function iconNote(s, x, y, w, h, c) {
    box(s, { x: x, y: y + h * 0.62, w: w * 0.52, h: h * 0.38, shape: 'ellipse', fill: c });
    box(s, { x: x + w * 0.44, y: y, w: w * 0.11, h: h * 0.78, fill: c });
    box(s, { x: x + w * 0.44, y: y, w: w * 0.56, h: h * 0.20, fill: c, shape: 'round2SameRect' });
}

function iconMic(s, x, y, w, h, c) {
    box(s, { x: x + w * 0.28, y: y, w: w * 0.44, h: h * 0.58, fill: c, rectRadius: 0.5, shape: 'roundRect' });
    box(s, { x: x + w * 0.08, y: y + h * 0.28, w: w * 0.84, h: h * 0.50, shape: 'arc', fill: null, line: { color: c, width: 2 }, angleRange: [0, 180] });
    box(s, { x: x + w * 0.45, y: y + h * 0.70, w: w * 0.10, h: h * 0.22, fill: c });
    box(s, { x: x + w * 0.18, y: y + h * 0.90, w: w * 0.64, h: h * 0.08, fill: c });
}

/* two ear-buds, each a filled cup with a cable curving away below it */
function iconHeadphones(s, x, y, w, h, c) {
    [0, 1].forEach(function (i) {
        const bx = x + w * (i ? 0.62 : 0);
        box(s, { x: bx, y: y, w: w * 0.38, h: h * 0.50, fill: c, shape: 'round2SameRect', rectRadius: 0.35, flipH: i === 1 });
        s.addShape('custGeom', {
            x: bx + w * (i ? 0.06 : 0.14), y: y + h * 0.44, w: w * 0.24, h: h * 0.56,
            fill: { type: 'none' }, line: { color: c, width: 2.5 },
            points: [
                { x: i ? w * 0.18 : w * 0.06, y: 0 },
                { x: i ? 0 : w * 0.24, y: h * 0.56, curve: { type: 'quadratic', x1: i ? w * 0.18 : w * 0.06, y1: h * 0.44 } },
            ],
        });
    });
}

/* turntable: rounded square frame, platter, spindle and tone-arm */
function iconDisc(s, x, y, w, h, c) {
    box(s, { x: x, y: y, w: w, h: h, fill: null, rectRadius: 0.12, shape: 'roundRect', line: { color: c, width: 3 } });
    box(s, { x: x + w * 0.14, y: y + h * 0.16, w: w * 0.58, h: h * 0.68, shape: 'ellipse', fill: c });
    box(s, { x: x + w * 0.38, y: y + h * 0.44, w: w * 0.10, h: h * 0.12, shape: 'ellipse', fill: W95 });
    box(s, { x: x + w * 0.68, y: y + h * 0.20, w: w * 0.08, h: h * 0.52, fill: c, rotate: 20, rectRadius: 0.4, shape: 'roundRect' });
}

/* head + shoulders with a small record in front */
function iconUser(s, x, y, w, h, c) {
    box(s, { x: x + w * 0.38, y: y, w: w * 0.36, h: h * 0.38, shape: 'ellipse', fill: c });
    box(s, { x: x + w * 0.24, y: y + h * 0.44, w: w * 0.74, h: h * 0.72, shape: 'ellipse', fill: c });
    box(s, { x: x, y: y + h * 0.52, w: w * 0.40, h: h * 0.42, shape: 'donut', fill: c, rectRadius: 0.3 });
}

function iconTreble(s, x, y, w, h, c) {
    box(s, { x: x + w * 0.36, y: y, w: w * 0.28, h: h * 0.78, shape: 'ellipse', fill: null, line: { color: c, width: 2 } });
    box(s, { x: x + w * 0.44, y: y + h * 0.10, w: w * 0.10, h: h * 0.80, fill: c });
    box(s, { x: x + w * 0.30, y: y + h * 0.72, w: w * 0.36, h: h * 0.28, shape: 'ellipse', fill: c });
}

/* ============================================================ slide builders */
const slides = [];

/* 1 - cover: giant SPOTLIGHT wordmark, ghosted above/below, between two rules */
slides.push(function (s) {
    s.background = { color: INK };
    tx(s, 'SPOTLIGHT', { x: 2.55, y: 2.31, w: 8.1, h: 1.15, align: 'center', fontFace: MSB, fontSize: 80, color: '494D50' });
    tx(s, 'SPOTLIGHT', { x: 2.54, y: 4.06, w: 7.84, h: 1.15, align: 'center', fontFace: MSB, fontSize: 80, color: '494D50' });
    s.addShape('line', { x: 2.51, y: 3.033, w: 8.313, h: 0, line: { color: W85, width: 3 } });
    s.addShape('line', { x: 2.51, y: 4.462, w: 8.313, h: 0, line: { color: W85, width: 3 } });
    tx(s, 'SPOTLIGHT', {
        x: 3.256, y: 3.033, w: 6.822, h: 1.447, align: 'center', valign: 'middle',
        fontFace: MSB, fontSize: 80, color: W,
    });
});

/* 2 - "Welcome to Spotlight Event Organizer" + magenta corner glow */
slides.push(function (s) {
    s.background = { color: BG };
    radialGlow(s, { cx: -0.2, cy: -0.2, r: 6.4, core: MAG, edge: BG, steps: 22 });
    tx(s, 'Welcome to Spotlight Event Organizer', { x: 1.194, y: 1.737, w: 3.501, h: 1.616, fontFace: PSB, fontSize: 30 });
    tx(s, 'About Our Event', { x: 1.194, y: 3.899, w: 1.562, h: 0.303, fontFace: PM, fontSize: 12, color: W95, wrap: false });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's.", {
        x: 1.198, y: 4.202, w: 3.501, h: 0.97, align: 'justify', fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5,
    });
    readMore(s, 1.299, 5.459);
    const card = rot90(5.482, 4.289, 2.269, 2.408);
    gradRect(s, { x: card.x, y: card.y, w: card.w, h: card.h, c1: '7E00BC', c2: '7B00BB', dir: 'a' });
    iconNote(s, 6.42, 4.79, 0.42, 0.72, W);
    tx(s, 'Our Music Journey', { x: 5.607, y: 5.911, w: 2.018, h: 0.337, align: 'center', fontFace: PM, fontSize: 14, color: W95, wrap: false });
});

/* 3 - "We Provide A Professional Music Organizer" */
slides.push(function (s) {
    s.background = { color: BG };
    box(s, { x: 1.04, y: 4.4, w: 2.195, h: 2.1, fill: PANEL, transparency: 42 });
    accentPanel(s, 6.543, 1.4, 6.79, 2.507); // layout's rotated accent bar
    tx(s, 'We Provide A Professional Music Organizer', { x: 7.693, y: 1.903, w: 4.493, h: 1.515, fontFace: PSB, fontSize: 28 });
    const cols = [
        { x: 4.48, w: 2.812, head: '15+ Years Experience', body: "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy." },
        { x: 8.49, w: 3.668, head: 'Professional Musician', body: "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s." },
    ];
    cols.forEach(function (c) {
        tx(s, c.head, { x: c.x, y: 4.641, w: 2.6, h: 0.303, fontFace: PM, fontSize: 12, color: W95, wrap: false });
        tx(s, c.body, { x: c.x, y: 4.977, w: c.w, h: 1.273, align: 'justify', fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5 });
    });
});

/* 4 - "Music is the Moonlight..." over a violet wash */
slides.push(function (s) {
    s.background = { color: BG };
    gradRect(s, { x: 0, y: 0, w: 13.344, h: 4.53, c1: '221237', c2: '5001A6', dir: 'd' });
    box(s, { x: 0, y: 3.42, w: 3.12, h: 1.11, fill: BG });
    tx(s, 'Music is the Moonlight in the Gloomy Night of Life.', { x: 6.532, y: 1.143, w: 5.654, h: 1.919, fontFace: PSB, fontSize: 36 });
    statCard(s, { x: 5.495, y: 3.957, w: 1.527, h: 1.323, value: '200K', vy: 0.361 });
    statCard(s, { x: 9.779, y: 3.957, w: 1.527, h: 1.323, value: '75%', vy: 0.361 });
    const feet = [
        { x: 4.635, label: 'Our Live Tour', body: 'Lorem Ipsum is simply dummy text of the printing typesetting industry. ' },
        { x: 8.919, label: 'Progress Project', body: 'Lorem Ipsum is simply dummy text of the printing typesetting industry.' },
    ];
    feet.forEach(function (f) {
        tx(s, f.label, { x: f.x, y: 5.671, w: 3.246, h: 0.404, align: 'center', fontFace: PM, fontSize: 18, color: W95, wrap: false });
        tx(s, f.body, { x: f.x, y: 6.089, w: 3.246, h: 0.667, align: 'center', fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5 });
    });
});

/* 5 - equipment gauges */
slides.push(function (s) {
    s.background = { color: BG };
    box(s, { x: 0, y: 1.339, w: 3.484, h: 6.161, fill: PANEL });
    box(s, { x: 3.302, y: 2.533, w: 3.127, h: 3.911, fill: PANEL });
    box(s, { x: 9.018, y: 4.536, w: 2.777, h: 2.964, fill: PANEL });
    box(s, { x: 6.241, y: 0, w: 7.092, h: 4.732, fill: PANEL });
    accentPanel(s, 3.302, 2.719, 2.957, 4.781);
    accentPanel(s, 9.021, 4.748, 2.5, 2.752);
    tx(s, 'Choose the Equipment That Fits Your Type', { x: 7.775, y: 0.924, w: 4.558, h: 1.969, fontFace: PSB, fontSize: 37 });
    gauge(s, { cx: 4.78, cy: 4.477, r: 0.85, pct: 80 });
    tx(s, 'Sound Mixing', { x: 3.836, y: 5.91, w: 1.889, h: 0.37, align: 'center', fontFace: PM, fontSize: 16, color: W95 });
    tx(s, 'Lorem Ipsum is simply dummy.', { x: 3.66, y: 6.27, w: 2.241, h: 0.342, align: 'center', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5 });
    gauge(s, { cx: 10.272, cy: 5.833, r: 0.64, pct: 70 });
    tx(s, 'Voice Record', { x: 9.327, y: 6.703, w: 1.889, h: 0.37, align: 'center', fontFace: PM, fontSize: 16, color: W95 });
});

/* 6 - "We Are Professional Musician" + 350+ badge */
slides.push(function (s) {
    s.background = { color: BG };
    const bar = rot90(9.444, 0.717, 3.713, 4.109);
    gradRect(s, { x: bar.x, y: bar.y, w: bar.w, h: bar.h, c1: '1D142F', c2: '5001AA', dir: 'd' });
    gradRect(s, { x: -0.016, y: -0.016, w: 2.411, h: 2.314, c1: 'A000BF', c2: '5300B4', dir: 'd' });
    tx(s, '350+', { x: 0.339, y: 0.535, w: 1.701, h: 0.841, align: 'center', valign: 'middle', fontFace: PSB, fontSize: 44, bold: true });
    tx(s, 'Our Project', { x: 0.527, y: 1.411, w: 1.325, h: 0.337, align: 'center', fontFace: PM, fontSize: 14, color: W95, wrap: false });
    tx(s, 'We Are Professional Musician', { x: 5.514, y: 5.175, w: 2.974, h: 1.515, fontFace: PSB, fontSize: 28 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s..", {
        x: 8.928, y: 5.345, w: 3.192, h: 1.175, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
    navBtn(s, { x: 10.841, y: 2.315, w: 0.987, h: 0.913, size: 28 });
});

/* 7 - musician card "Erick Moore Koprai" */
slides.push(function (s) {
    s.background = { color: BG };
    const card = rot90(2.453, 0.52, 4.087, 6.46);
    gradRect(s, { x: card.x, y: card.y, w: card.w, h: card.h, c1: '67087F', c2: '5200B1', dir: 'v' });
    tx(s, 'Erick Moore Koprai', { x: 2.115, y: 2.336, w: 4.763, h: 1.447, fontFace: PSB, fontSize: 40, color: W95 });
    tx(s, 'Young Musician', { x: 2.115, y: 4.111, w: 2.6, h: 0.37, fontFace: PM, fontSize: 16, color: W85, wrap: false });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's.", {
        x: 2.115, y: 4.545, w: 4.096, h: 0.62, fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
    navBtn(s, { x: 9.266, y: 5.664, w: 0.485, h: 0.523, ch: '<', fill: BG, color: W75, size: 14 });
    navBtn(s, { x: 9.856, y: 5.64, w: 0.544, h: 0.585, size: 16 });
});

/* 8 - "Our Amazing Team" grid over nested chevrons */
slides.push(function (s) {
    s.background = { color: BG };
    /* three nested ">" chevrons: two hairline outlines then the solid one */
    const chevron = function (dx, fill, line) {
        s.addShape('custGeom', {
            x: 7.09 - dx, y: 0, w: 5.93, h: 7.5,
            fill: fill ? { color: fill } : { type: 'none' },
            line: line || { type: 'none' },
            points: [
                { x: 0.48, y: 0 }, { x: 2.58, y: 0 }, { x: 5.93, y: 3.5 },
                { x: 2.10, y: 7.5 }, { x: 0, y: 7.5 }, { x: 3.83, y: 3.5 }, { close: true },
            ],
        });
    };
    chevron(1.64, null, { color: VIO, width: 1 });
    chevron(0.79, null, { color: MAG, width: 1 });
    chevron(0, '9C00C0');
    gradRect(s, { x: 5.874, y: 1.996, w: 1.851, h: 1.64, c1: '4A4757', c2: MAG, dir: 'v' });
    tx(s, 'Nicolas McJagger', { x: 5.874, y: 2.989, w: 1.851, h: 0.286, align: 'center', fontFace: PM, fontSize: 11, wrap: false });
    tx(s, 'Producer / DJ', { x: 5.874, y: 3.275, w: 1.851, h: 0.278, align: 'center', fontFace: PL, fontSize: 10.5, wrap: false });
    tx(s, 'Our Amazing Team', { x: 1.184, y: 2.259, w: 2.964, h: 1.043, fontFace: PSB, fontSize: 28, charSpacing: 0.5 });
    tx(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum.', {
        x: 1.162, y: 3.574, w: 2.816, h: 0.897, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
    readMore(s, 1.279, 4.937);
});

/* 9 - "Meet Enzi Centurio" profile */
slides.push(function (s) {
    s.background = { color: BG };
    radialGlow(s, { cx: 13.6, cy: 7.7, r: 6.6, core: MAG, edge: BG, steps: 22 });
    box(s, { x: 0, y: 5.168, w: 5.289, h: 1.624, fill: BG, line: { color: '2E2C46', width: 1 } });
    tx(s, 'Enzi Centurio', { x: 1.148, y: 5.703, w: 3.2, h: 0.572, fontFace: PSB, fontSize: 28, charSpacing: 1, wrap: false });
    tx(s, 'Expert Composer', { x: 1.156, y: 6.268, w: 2.6, h: 0.404, fontFace: PL, fontSize: 18, color: W75, wrap: false });
    navBtn(s, { x: 5.309, y: 3.327, w: 1.085, h: 1.003, size: 30 });
    tx(s, 'Meet ', { x: 7.785, y: 1.38, w: 3.9, h: 0.707, fontFace: PSB, fontSize: 36 });
    tx(s, 'Enzi Centurio', { x: 7.778, y: 2.091, w: 3.919, h: 0.707, fontFace: PSB, fontSize: 36, charSpacing: 1 });
    [3.013, 4.378].forEach(function (y) {
        gradRect(s, { x: 7.866, y: y, w: 4.153, h: 0.03, c1: MAG, c2: BG, steps: 16 });
    });
    tx(s, 'Education & Experience', { x: 7.784, y: 3.191, w: 3.0, h: 0.337, fontFace: PM, fontSize: 14, color: W95, wrap: false });
    const rows = [
        ['Degree at Music University', 7.795, 3.52, 2.374], ['Head of Spotlight Studio', 7.795, 3.826, 2.374],
        ['Musician', 10.938, 3.527, 1.071], ['Composer', 10.938, 3.842, 0.949],
    ];
    rows.forEach(function (r) {
        tx(s, r[0], { x: r[1], y: r[2], w: r[3], h: 0.342, fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5 });
    });
    tx(s, 'About Our Composer', { x: 7.779, y: 4.565, w: 2.6, h: 0.337, fontFace: PM, fontSize: 14, color: W95, wrap: false });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled.", {
        x: 7.779, y: 4.945, w: 4.332, h: 1.175, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
});

/* 10 - "Our Main Service" with two device mock-ups */
slides.push(function (s) {
    s.background = { color: BG };
    tx(s, 'Our Main Service', { x: 1.169, y: 5.069, w: 2.976, h: 0.438, fontFace: PSB, fontSize: 20 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard.", {
        x: 1.165, y: 5.506, w: 3.57, h: 0.97, align: 'justify', fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5,
    });
    iconDisc(s, 5.442, 5.399, 0.76, 0.736, MAG);
    tx(s, 'Multiple Studios', { x: 6.487, y: 5.402, w: 2.2, h: 0.32, fontFace: PM, fontSize: 13, color: W95, wrap: false });
    tx(s, 'Lorem Ipsum is simply dummy', { x: 6.497, y: 5.722, w: 2.237, h: 0.342, fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5 });
    iconUser(s, 9.014, 5.399, 0.657, 0.62, MAG);
    tx(s, 'Live Instrument', { x: 9.953, y: 5.402, w: 2.2, h: 0.32, fontFace: PM, fontSize: 13, color: W95, wrap: false });
    tx(s, 'Lorem Ipsum is simply dummy', { x: 9.963, y: 5.722, w: 2.247, h: 0.342, fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5 });
});

/* 11 - "We Have Amazing Jobs All Over The World" arrow panel */
slides.push(function (s) {
    s.background = { color: BG };
    gradRect(s, { x: 4.072, y: 0, w: 4.615, h: 7.5, c1: MAG, c2: VIO, dir: 'v' });
    tx(s, 'We Have Amazing Jobs All Over The World', {
        x: 4.725, y: 2.077, w: 3.052, h: 1.986, align: 'center', fontFace: PSB, fontSize: 28,
    });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever.", {
        x: 4.837, y: 4.265, w: 2.829, h: 1.175, align: 'center', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
    navBtn(s, { x: 8.33, y: 3.396, w: 0.711, h: 0.707, size: 18 });
    iconHeadphones(s, 10.624, 1.278, 0.708, 0.708, VIO);
    tx(s, 'Best Equipment', { x: 9.524, y: 2.225, w: 2.942, h: 0.37, align: 'center', fontFace: PM, fontSize: 16, color: W95 });
    tx(s, 'Lorem Ipsum is simply dummy text of the printing.', { x: 9.817, y: 2.632, w: 2.413, h: 0.62, align: 'center', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5 });
    iconMic(s, 10.727, 4.114, 0.5, 0.917, VIO);
    tx(s, 'Audio Quality', { x: 9.507, y: 5.272, w: 2.942, h: 0.37, align: 'center', fontFace: PM, fontSize: 16, color: W95 });
    tx(s, 'Lorem Ipsum is simply dummy text of the printing.', { x: 9.8, y: 5.679, w: 2.413, h: 0.62, align: 'center', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5 });
});

/* 12 - "Our Amazing Services" - four outlined cards, first one filled */
slides.push(function (s) {
    s.background = { color: BG };
    tx(s, 'Our Amazing Services', { x: 3.538, y: 1.029, w: 6.256, h: 0.572, align: 'center', fontFace: PSB, fontSize: 28 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever.", {
        x: 3.444, y: 1.657, w: 6.445, h: 0.667, align: 'center', fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5,
    });
    const cards = [
        { x: 1.015, title: 'Live Instrument', icon: iconUser, ix: 1.953, iw: 0.657, ih: 0.62, filled: true },
        { x: 4.0, title: 'Audio Quality', icon: iconMic, ix: 5.091, iw: 0.401, ih: 0.736, filled: false },
        { x: 6.917, title: 'Best Equipment', icon: iconHeadphones, ix: 7.787, iw: 0.763, ih: 0.7, filled: false },
        { x: 9.833, title: 'Multiple Studios', icon: iconDisc, ix: 10.625, iw: 0.76, ih: 0.736, filled: false },
    ];
    cards.forEach(function (c, i) {
        if (c.filled) accentPanel(s, c.x, 2.994, 2.485, 3.433, 'd');
        else box(s, { x: c.x, y: 2.994, w: 2.485, h: 3.433, fill: null, line: { color: '7B3FBE', width: 1 } });
        c.icon(s, c.ix, 3.71, c.iw, c.ih, c.filled ? W : MAG);
        tx(s, c.title, { x: c.x, y: 4.922, w: 2.485, h: 0.353, align: 'center', fontFace: PM, fontSize: 15, color: W95, wrap: false });
        tx(s, 'Lorem Ipsum is simply dummy text of the printing.', {
            x: c.x + 0.02, y: 5.301, w: 2.445, h: 0.62, align: 'center', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
        });
    });
});

/* 13 - section break */
slides.push(function (s) {
    s.background = { color: INK };
    tx(s, 'It’s time to ', { x: 7.101, y: 1.759, w: 2.4, h: 0.558, fontFace: PM, fontSize: 20, color: W95, lineSpacingMultiple: 1.5 });
    tx(s, 'SECTION BREAK', { x: 7.06, y: 2.302, w: 5.438, h: 0.825, fontFace: MSB, fontSize: 43, color: W95 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown.  ", {
        x: 7.079, y: 3.359, w: 5.316, h: 0.897, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
    box(s, { x: 7.172, y: 5.066, w: 3.0, h: 0.675, fill: W95, rectRadius: 0.12 });
    s.addText([
        { text: '12.00 PM ', options: { color: MAG } },
        { text: '- 12.15 PM', options: { color: VIO } },
    ], { x: 7.172, y: 5.066, w: 3.0, h: 0.675, align: 'center', valign: 'middle', fontFace: PM, fontSize: 18 });
});

/* 14 - "Music Sound Different..." photo mosaic */
slides.push(function (s) {
    s.background = { color: BG };
    tx(s, 'Music Sound Different to the One Who Plays it.', { x: 1.152, y: 1.547, w: 4.316, h: 0.909, fontFace: PSB, fontSize: 24 });
    tx(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the.', {
        x: 1.157, y: 2.684, w: 3.898, h: 0.62, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
    tx(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been.', {
        x: 9.556, y: 4.181, w: 2.584, h: 0.897, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
    gradRect(s, { x: 7.936, y: 5.517, w: 4.292, h: 0.924, c1: VIO, c2: MAG, dir: 'd' });
    iconNote(s, 8.366, 5.758, 0.254, 0.429, W);
    s.addText([
        { text: 'Lorem Ipsum ', options: { fontFace: PM } },
        { text: 'is simply dummy', options: { fontFace: PL } },
    ], { x: 8.783, y: 5.811, w: 3.207, h: 0.337, fontSize: 14, color: W95, isTextBox: true });
});

/* 15 - "Vocal Mastering" numbered tiles */
slides.push(function (s) {
    s.background = { color: BG };
    box(s, { x: 0, y: 0, w: 6.003, h: 7.498, fill: PANEL });
    /* the three unfilled photo wells read as slightly lighter squares */
    [[0.883, 1.095, 1.846, 1.769], [0.883, 2.865, 1.846, 1.769], [4.364, 4.641, 1.636, 1.754]].forEach(function (r) {
        box(s, { x: r[0], y: r[1], w: r[2], h: r[3], fill: '2A2840' });
    });
    accentPanel(s, 2.728, 2.865, 1.636, 1.769, 'd');
    tx(s, 'High Quality Audio Recording Into Record', { x: 3.489, y: 1.575, w: 3.821, h: 0.808, fontFace: PM, fontSize: 21, color: W95 });
    const tiles = [
        ['01', 'Responsive', 2.84, W], ['02', 'Innovative', 4.476, W95],
    ];
    tiles.forEach(function (t) {
        tx(s, t[0], { x: t[2], y: 3.207, w: 1.2, h: 0.774, fontFace: PSB, fontSize: 40, color: t[3] });
        tx(s, t[1], { x: t[2], y: 3.99, w: 1.4, h: 0.303, fontFace: PM, fontSize: 12, color: W95, wrap: false });
    });
    tx(s, '03', { x: 2.84, y: 4.976, w: 1.2, h: 0.774, fontFace: PSB, fontSize: 40, color: W95 });
    tx(s, 'Passionate', { x: 2.84, y: 5.759, w: 1.4, h: 0.303, fontFace: PM, fontSize: 12, color: W95, wrap: false });
    navBtn(s, { x: 1.38, y: 5.129, w: 0.852, h: 0.909, ch: '<', size: 26 });
    tx(s, 'Vocal Mastering With Various Type', { x: 8.469, y: 1.795, w: 3.468, h: 2.524, fontFace: PSB, fontSize: 36 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an.", {
        x: 8.469, y: 4.53, w: 3.336, h: 1.175, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
});

/* 16 - "Enjoy Chicago-Style Rock Music..." */
slides.push(function (s) {
    s.background = { color: BG };
    tx(s, 'Enjoy Chicago-Style Rock Music From The Energetic Frank Cheese', { x: 7.334, y: 1.071, w: 4.85, h: 1.313, fontFace: PSB, fontSize: 24 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen book.", {
        x: 1.157, y: 4.095, w: 3.084, h: 1.731, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
    [['The Ultimate Duo Tribute', 5.276], ['Holiday Modern Concert', 8.891]].forEach(function (c) {
        tx(s, c[0], { x: c[1], y: 5.789, w: 2.973, h: 0.337, align: 'center', fontFace: PM, fontSize: 14, color: W95, wrap: false });
        tx(s, 'Lorem Ipsum is simply.', { x: c[1], y: 6.119, w: 2.973, h: 0.342, align: 'center', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5 });
    });
});

/* 17 - "The Countdown to Thursday Afternoon's Music City Bowl" */
slides.push(function (s) {
    s.background = { color: BG };
    gradRect(s, { x: 1.265, y: 3.939, w: 2.499, h: 2.601, c1: '40058E', c2: '4D00AE', dir: 'v' });
    iconNote(s, 2.295, 4.352, 0.439, 0.742, W);
    tx(s, 'Outdoor Music', { x: 1.657, y: 5.29, w: 1.716, h: 0.337, align: 'center', fontFace: PM, fontSize: 14, color: W95 });
    tx(s, 'Lorem Ipsum is simply dummy printing typesetting.', {
        x: 1.409, y: 5.643, w: 2.212, h: 0.573, align: 'center', fontFace: LATO, fontSize: 10, color: W75, lineSpacingMultiple: 1.5,
    });
    tx(s, "The Countdown to Thursday Afternoon's Music City Bowl", { x: 8.199, y: 1.696, w: 3.947, h: 1.313, fontFace: PSB, fontSize: 24 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen book.", {
        x: 8.199, y: 3.311, w: 3.946, h: 1.453, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
    tx(s, 'Play A Concert At Winter', { x: 8.199, y: 5.45, w: 3.0, h: 0.353, fontFace: PL, fontSize: 15, color: W95, wrap: false });
    s.addShape('line', { x: 11.251, y: 5.629, w: 0.768, h: 0, line: { color: MAG, width: 2.75 } });
});

/* 18 - "The Spotlight Will Offer A Show of Jazz..." icon rail */
slides.push(function (s) {
    s.background = { color: BG };
    accentPanel(s, 5.627, 1.267, 1.356, 4.955, 'd');
    const rail = [
        { icon: iconHeadphones, x: 6.0, y: 1.57, w: 0.61, h: 0.61, label: 'Equipment', ly: 2.351 },
        { icon: iconMic, x: 6.105, y: 3.148, w: 0.401, h: 0.736, label: 'Audio', ly: 4.109 },
        { icon: iconDisc, x: 5.918, y: 5.005, w: 0.764, h: 0.639, label: 'Multiple', ly: 5.809 },
    ];
    rail.forEach(function (r) {
        r.icon(s, r.x, r.y, r.w, r.h, W);
        tx(s, r.label, { x: 5.729, y: r.ly, w: 1.153, h: 0.303, align: 'center', fontFace: PR, fontSize: 12, color: W95 });
    });
    tx(s, 'The Spotlight Will Offer A Show of Jazz, Blues, Soul, R&B, and Rock n’ Roll', { x: 7.65, y: 1.825, w: 4.583, h: 1.986, fontFace: PSB, fontSize: 28 });
    tx(s, 'Brightest Public Displays', { x: 7.658, y: 4.176, w: 3.259, h: 0.337, fontFace: PM, fontSize: 14, color: W95 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type.", {
        x: 7.65, y: 4.5, w: 4.491, h: 1.175, align: 'justify', fontFace: LATO, fontSize: 11, color: GREY50, lineSpacingMultiple: 1.5,
    });
});

/* 19 - "Special Mockup Music" laptop mock-ups inside concentric rings */
slides.push(function (s) {
    s.background = { color: BG };
    [2.7, 2.35, 2.0].forEach(function (r) {
        box(s, { x: 6.666 - r, y: 1.75 - r, w: r * 2, h: r * 2, shape: 'ellipse', fill: null, line: { color: '8E1BAE', width: 1 } });
    });
    gradEllipse(s, { x: 3.86, y: -0.85, w: 5.4, h: 5.25, c1: '9600C0', c2: '6E00BA' });
    tx(s, 'Special Mockup Music', { x: 4.655, y: 1.379, w: 3.809, h: 0.976, align: 'center', fontFace: PSB, fontSize: 26 });
    imgBox(s, 1.678, 3.046, 3.475, 1.913, '[image] laptop');
    imgBox(s, 7.95, 3.046, 3.475, 1.913, '[image] laptop');
    imgBox(s, 4.138, 3.235, 4.876, 2.617, '[image] laptop');
    s.addText([
        { text: '“Lorem Ipsum ', options: { bold: true, italic: true, color: W95 } },
        { text: "is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard.”", options: { italic: true, color: W75 } },
    ], { x: 3.786, y: 6.359, w: 5.762, h: 0.62, align: 'center', fontFace: LATO, fontSize: 11, lineSpacingMultiple: 1.5, isTextBox: true });
});

/* 20 - "Amazing Events Touring By SpotLight Production" + phone mock-up */
slides.push(function (s) {
    s.background = { color: BG };
    accentPanel(s, 10.023, 0, 3.311, 7.517, 'd');
    imgBox(s, 8.293, 1.562, 3.604, 5.938, '[image] phone');
    tx(s, 'Amazing Events Touring By SpotLight Production', { x: 3.349, y: 1.357, w: 3.994, h: 1.313, fontFace: PSB, fontSize: 24 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer took a galley of type and scrambled it to make a type specimen.", {
        x: 3.339, y: 2.85, w: 3.994, h: 1.576, align: 'justify', fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5,
    });
    statCard(s, { x: 3.45, y: 5.056, w: 1.666, h: 1.087, value: '4530', vsize: 24, vy: 0.18, vcolor: MAG, label: 'Successful Event', ly: 0.637 });
    statCard(s, { x: 5.573, y: 5.056, w: 1.666, h: 1.087, value: '300+', vsize: 24, vy: 0.18, vcolor: MAG, label: 'Professional Team', ly: 0.637 });
});

/* 21 - stats rail + monitor mock-up */
slides.push(function (s) {
    s.background = { color: BG };
    accentPanel(s, 0.5, 0, 4.362, 7.5);
    const stats = [['800+', 'User Favorite', 0.919], ['1520', 'Music Project', 3.163], ['2512', 'Studio Room', 5.422]];
    stats.forEach(function (st) {
        tx(s, st[0], { x: 1.1, y: st[2], w: 2.0, h: 0.909, align: 'center', fontFace: PM, fontSize: 48 });
        tx(s, st[1], { x: 1.1, y: st[2] + 0.891, w: 2.0, h: 0.303, align: 'center', fontFace: PM, fontSize: 12, color: W95, wrap: false });
    });
    imgBox(s, 3.359, 2.097, 4.774, 3.867, '[image] monitor');
    tx(s, 'You Can Make A Reservation Through Offical Website', { x: 9.18, y: 1.315, w: 3.131, h: 3.332, fontFace: PSB, fontSize: 32 });
    tx(s, 'Our Great Mockup', { x: 9.186, y: 4.975, w: 2.5, h: 0.337, fontFace: PM, fontSize: 14, color: W95, wrap: false });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's.", {
        x: 9.211, y: 5.288, w: 2.899, h: 0.897, align: 'justify', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
    });
});

/* 22 - diamond chain infographic */
slides.push(function (s) {
    s.background = { color: BG };
    tx(s, 'Infographic Mockup Slide', { x: 1.158, y: 1.002, w: 3.887, h: 1.313, fontFace: PSB, fontSize: 36 });
    const nodes = [
        { x: 3.495, y: 4.197, icon: iconMic, ix: 4.129, iy: 4.65, iw: 0.357, ih: 0.654 },
        { x: 5.056, y: 2.612, icon: iconDisc, ix: 5.539, iy: 3.151, iw: 0.6, ih: 0.559 },
        { x: 6.574, y: 4.151, icon: iconUser, ix: 7.164, iy: 4.625, iw: 0.584, ih: 0.6 },
        { x: 8.178, y: 2.612, icon: iconHeadphones, ix: 8.74, iy: 3.179, iw: 0.525, ih: 0.525 },
    ];
    nodes.forEach(function (n) {
        box(s, { x: n.x, y: n.y, w: 1.72, h: 1.61, shape: 'diamond', fill: MAG });
        box(s, { x: n.x + 0.29, y: n.y + 0.23, w: 1.13, h: 1.18, shape: 'ellipse', fill: W95 });
        n.icon(s, n.ix, n.iy, n.iw, n.ih, VIO);
    });
    [[3.118, 5.0, 0.223, 0.55], [10.035, 3.4, 0.6, 0.06], [7.35, 3.5, 0.06, 0.7]].forEach(function (c) {
        box(s, { x: c[0], y: c[1], w: c[2], h: c[3], fill: VIO });
    });
    const labels = [
        { t: 'Audio', x: 1.244, ty: 4.455, by: 4.829 },
        { t: 'Instrument', x: 6.719, ty: 1.606, by: 1.979 },
        { t: 'Studios', x: 5.227, ty: 5.422, by: 5.796 },
        { t: 'Equipment', x: 10.619, ty: 2.885, by: 3.259 },
    ];
    labels.forEach(function (l) {
        tx(s, l.t, { x: l.x, y: l.ty, w: 1.446, h: 0.353, align: 'center', valign: 'bottom', fontFace: PM, fontSize: 15, color: W95 });
        tx(s, 'Lorem Ipsum is simply dummy text.', { x: l.x, y: l.by, w: 1.47, h: 0.669, align: 'center', fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5 });
    });
    tx(s, 'Lorem Ipsum is simply dummy the printing and typesetting industry.  Lorem Ipsum has.', {
        x: 9.517, y: 4.854, w: 2.89, h: 0.97, fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5,
    });
    readMore(s, 9.604, 6.112);
});

/* 23 - three linked circles infographic */
slides.push(function (s) {
    s.background = { color: BG };
    tx(s, 'Infographic Mockup Slide', { x: 3.698, y: 1.002, w: 5.937, h: 0.64, align: 'center', fontFace: PSB, fontSize: 32 });
    const items = [
        { cx: 2.284, num: '01', icon: iconHeadphones, ix: 2.949, iy: 3.272, iw: 1.048, ih: 1.0, label: 'Equipment', lx: 2.302 },
        { cx: 5.509, num: '02', icon: iconMic, ix: 6.385, iy: 3.226, iw: 0.6, ih: 1.1, label: 'Audio', lx: 5.588 },
        { cx: 8.875, num: '03', icon: iconDisc, ix: 9.512, iy: 3.268, iw: 1.0, ih: 1.056, label: 'Instrument', lx: 8.66 },
    ];
    items.forEach(function (it, i) {
        box(s, { x: it.cx, y: 2.592, w: 2.447, h: 2.447, shape: 'ellipse', fill: MAG });
        box(s, { x: it.cx + 0.5, y: 2.592, w: 1.947, h: 2.447, shape: 'pie', fill: VIO, angleRange: [270, 90] });
        it.icon(s, it.ix, it.iy, it.iw, it.ih, W);
        box(s, { x: it.cx - 0.474, y: 3.384, w: 0.863, h: 0.863, shape: 'ellipse', fill: W });
        tx(s, it.num, {
            x: it.cx - 0.474, y: 3.384, w: 0.863, h: 0.863, align: 'center', valign: 'middle',
            fontFace: PM, fontSize: 20, italic: true, color: GREY25,
        });
        tx(s, it.label, { x: it.lx, y: 5.4, w: 2.138, h: 0.404, align: 'center', valign: 'bottom', fontFace: PM, fontSize: 18, color: W95 });
        tx(s, 'Lorem Ipsum is simply dummy text of the printing', {
            x: i === 2 ? 8.948 : it.lx, y: 5.817, w: 2.138, h: 0.669, align: 'center', fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5,
        });
    });
});

/* 24 - "planet" infographic (sphere quarters + ring) */
slides.push(function (s) {
    s.background = { color: BG };
    tx(s, 'Infographic Mockup Slide', { x: 1.636, y: 1.002, w: 6.158, h: 0.64, align: 'center', fontFace: PSB, fontSize: 32 });
    gradEllipse(s, { x: 4.493, y: 2.565, w: 3.852, h: 3.854, c1: '6C00B9', c2: '9F00C1' });
    box(s, { x: 6.41, y: 2.565, w: 0.03, h: 3.854, fill: BG });
    box(s, { x: 4.493, y: 4.397, w: 3.852, h: 0.03, fill: BG });
    box(s, { x: 3.283, y: 3.683, w: 6.268, h: 1.787, shape: 'ellipse', fill: W85, line: { color: MAG, width: 2 } });
    box(s, { x: 4.05, y: 4.05, w: 4.75, h: 1.06, shape: 'ellipse', fill: BG });
    const glyphs = [
        { icon: iconMic, x: 5.244, y: 3.008, w: 0.28, h: 0.5 },
        { icon: iconDisc, x: 6.753, y: 3.015, w: 0.435, h: 0.4 },
        { icon: iconUser, x: 5.244, y: 4.079, w: 0.346, h: 0.36 },
        { icon: iconTreble, x: 7.249, y: 4.02, w: 0.2, h: 0.47 },
        { icon: iconNote, x: 5.615, y: 5.6, w: 0.32, h: 0.4 },
        { icon: iconHeadphones, x: 6.888, y: 5.63, w: 0.32, h: 0.36 },
    ];
    glyphs.forEach(function (g) { g.icon(s, g.x, g.y, g.w, g.h, W); });
    [['Creative Project', 2.433, 2.945], ['Reliable Music', 4.87, 5.383]].forEach(function (c) {
        tx(s, c[0], { x: 9.477, y: c[1], w: 2.672, h: 0.512, fontFace: PSB, fontSize: 18, color: W95, lineSpacingMultiple: 1.5 });
        tx(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has.', {
            x: 9.477, y: c[2], w: 2.712, h: 0.897, fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
        });
    });
});

/* 25 - pie chart + magenta copy panel */
slides.push(function (s) {
    s.background = { color: BG };
    s.addChart('pie', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'], values: [8.2, 3.2, 1.4, 1.2] }], {
        x: 0.011, y: 1.401, w: 7.047, h: 4.698,
        chartColors: [VIO, DEEP, MAG, W95],
        dataBorder: { pt: 0.75, color: W },
        showLegend: true, legendPos: 'b', legendColor: W, legendFontFace: PM, legendFontSize: 12,
        showTitle: false, chartArea: { fill: { color: BG } }, plotArea: { fill: { color: BG } },
    });
    accentPanel(s, 6.777, 0.797, 5.262, 6.703, 'd');
    tx(s, 'We Have Amazing Jobs All Over The World', { x: 7.388, y: 1.808, w: 4.041, h: 1.515, fontFace: PSB, fontSize: 28 });
    tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard dummy text ever since the 1500s,  when an unknown printer took a galley of type", {
        x: 7.372, y: 3.678, w: 4.072, h: 1.175, align: 'justify', fontFace: LATO, fontSize: 11, color: W95, lineSpacingMultiple: 1.5,
    });
    [['4500', 'Electronic', 7.372], ['2900', 'R & B', 8.764], ['5300', 'Pop ', 10.157]].forEach(function (c) {
        statCard(s, { x: c[2], y: 5.388, w: 1.287, h: 1.1, value: c[0], vsize: 27, vy: 0.15, vcolor: W, label: c[1], ly: 0.682 });
    });
});

/* 26 - 3-D bar chart + "Sales Analysis Slides" */
slides.push(function (s) {
    s.background = { color: BG };
    s.addChart('bar3d', [
        { name: 'Series 1', labels: ['Audio', 'Studio', 'Equipment', 'Instrument'], values: [4.3, 2.5, 3.5, 4.5] },
        { name: 'Series 2', labels: ['Audio', 'Studio', 'Equipment', 'Instrument'], values: [2.4, 4.4, 1.8, 2.8] },
    ], {
        x: 0.948, y: 0.858, w: 6.68, h: 5.784,
        barDir: 'bar', barGrouping: 'clustered', bar3DShape: 'box',
        chartColors: [VIO, W],
        showLegend: false, showTitle: false,
        chartArea: { fill: { color: BG } }, plotArea: { fill: { color: BG } },
        catAxisLabelColor: W95, valAxisLabelColor: W95,
        catAxisLabelFontFace: PM, valAxisLabelFontFace: PM,
        catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
        catAxisLineColor: '3A3852', valGridLine: { color: '3A3852', style: 'solid', size: 0.75 },
    });
    tx(s, 'Sales Analysis Slides', { x: 8.397, y: 1.265, w: 3.815, h: 1.313, fontFace: PSB, fontSize: 36 });
    [['Project Overview', 3.052, 3.481], ['Music Report', 4.912, 5.341]].forEach(function (c) {
        tx(s, c[0], { x: 8.506, y: c[1], w: 2.427, h: 0.404, valign: 'bottom', fontFace: PM, fontSize: 18, color: W95 });
        tx(s, "Lorem Ipsum is simply dummy text of the printing and typesetting industry.  Lorem Ipsum has been the industry's standard.", {
            x: 8.506, y: c[2], w: 3.526, h: 0.97, align: 'justify', fontFace: LATO, fontSize: 12, color: W75, lineSpacingMultiple: 1.5,
        });
    });
});

/* 27 - "Months of Good Music" pricing table */
slides.push(function (s) {
    s.background = { color: BG };
    tx(s, 'Months of Good Music', { x: 3.255, y: 0.942, w: 6.823, h: 0.707, align: 'center', fontFace: PSB, fontSize: 36 });
    const plans = [
        { x: 1.269, title: 'SILVER PLAN', price: '$60,', feats: ['First Feature', 'Second Feature'], hero: false },
        { x: 5.047, title: 'GOLD PLAN', price: '$85,', feats: ['First Feature', 'Second Feature', 'Third Feature'], hero: true },
        { x: 8.848, title: 'PLATINUM PLAN', price: '$95,', feats: ['First Feature', 'Second Feature'], hero: false },
    ];
    plans.forEach(function (p) {
        const dy = p.hero ? -0.01 : 0;
        if (p.hero) {
            box(s, { x: p.x, y: 2.338, w: 3.212, h: 4.135, fill: PANEL, line: { color: MAG, width: 1 } });
            gradRect(s, { x: 5.037, y: 1.952, w: 3.231, h: 0.393, c1: '791273', c2: VIO });
            tx(s, 'Most Popular', { x: 5.651, y: 2.03, w: 2.03, h: 0.375, align: 'center', fontFace: PM, fontSize: 12, charSpacing: 3 });
        } else {
            box(s, { x: p.x, y: 2.548, w: 3.187, h: 3.854, fill: PANEL });
        }
        const cx = p.hero ? 5.61 : p.x + 0.558;
        tx(s, p.title, { x: cx, y: 2.74 + (p.hero ? -0.011 : 0), w: 2.072, h: 0.421, align: 'center', fontFace: PM, fontSize: 14, color: W95, lineSpacingMultiple: 1.5 });
        tx(s, 'Lorem Ipsum is simply dummy text of the printing and industry.', {
            x: cx - 0.271, y: 3.245 + dy, w: 2.614, h: 0.62, align: 'center', fontFace: LATO, fontSize: 11, color: W75, lineSpacingMultiple: 1.5,
        });
        s.addText([
            { text: p.price, options: { fontSize: 36 } },
            { text: '00', options: { fontSize: 28 } },
        ], { x: cx, y: 3.834 + dy, w: 2.072, h: 0.924, align: 'center', fontFace: PM, color: W, lineSpacingMultiple: 1.5, isTextBox: true });
        tx(s, 'FEATURES', { x: cx - 0.005, y: 4.757 + dy, w: 2.072, h: 0.421, align: 'center', fontFace: PM, fontSize: 14, color: W85, lineSpacingMultiple: 1.5 });
        p.feats.forEach(function (f, i) {
            const fy = 5.438 + dy + i * 0.397;
            box(s, { x: cx - 0.007, y: fy, w: 0.234, h: 0.234, shape: 'ellipse', fill: MAG });
            tx(s, f, { x: cx + 0.301, y: fy - 0.1, w: 1.8, h: 0.375, fontFace: PM, fontSize: 12, color: W85, lineSpacingMultiple: 1.5 });
        });
    });
});

/* 28 - contact banner */
slides.push(function (s) {
    s.background = { color: BG };
    accentPanel(s, 1.269, 2.875, 10.795, 3.021, 'd');
    box(s, { x: 2.209, y: 5.248, w: 8.915, h: 1.357, fill: BG });
    tx(s, 'Let’s Keep in Touch with Us? ', { x: 1.763, y: 3.482, w: 3.231, h: 1.043, fontFace: PSB, fontSize: 28 });
    box(s, { x: 5.875, y: 3.539, w: 0.173, h: 0.264, shape: 'teardrop', fill: W, rotate: 225 });
    tx(s, 'Our Address', { x: 6.234, y: 3.459, w: 2.0, h: 0.353, fontFace: PM, fontSize: 15, color: W95, charSpacing: 0.5, wrap: false });
    tx(s, '1234 Budiono Street , IN 012  Soerakarta, Indonesia', { x: 6.233, y: 3.822, w: 2.422, h: 0.667, fontFace: LATO, fontSize: 12, color: W85, lineSpacingMultiple: 1.5 });
    box(s, { x: 9.091, y: 3.53, w: 0.279, h: 0.279, shape: 'ellipse', fill: W });
    tx(s, 'Phone & Email', { x: 9.503, y: 3.459, w: 2.0, h: 0.353, fontFace: PM, fontSize: 15, color: W95, charSpacing: 0.5, wrap: false });
    tx(s, '(0123) 4567 - 8910\nYouremail@music.com', { x: 9.502, y: 3.822, w: 2.068, h: 0.667, fontFace: LATO, fontSize: 12, color: W85, lineSpacingMultiple: 1.5 });
    tx(s, 'Lorem Ipsum is simply dummy text of the printing and typesetting.', {
        x: 2.512, y: 5.587, w: 3.116, h: 0.678, align: 'justify', fontFace: PR, fontSize: 12, color: W75, lineSpacingMultiple: 1.5,
    });
    box(s, { x: 6.117, y: 5.623, w: 2.135, h: 0.591, fill: W });
    tx(s, 'MON-FRI: 9AM – 5PM', { x: 6.117, y: 5.623, w: 2.135, h: 0.591, align: 'center', valign: 'middle', fontFace: PM, fontSize: 10, color: DEEP, charSpacing: 0.5 });
    accentPanel(s, 8.248, 5.623, 2.135, 0.591, 'd');
    tx(s, 'SAT-SUN: 10AM – 20PM', { x: 8.248, y: 5.623, w: 2.135, h: 0.591, align: 'center', valign: 'middle', fontFace: PM, fontSize: 10, color: W95, charSpacing: 0.5 });
});

/* 29 - quote slide */
slides.push(function (s) {
    s.background = { color: BG };
    radialGlow(s, { cx: 1.6, cy: 0.4, r: 5.4, core: MAG, edge: BG, steps: 20 });
    gradRect(s, { x: 3.06, y: 2.105, w: 8.949, h: 3.25, c1: VIO, c2: MAG, dir: 'd' });
    box(s, { x: 3.9, y: 1.581, w: 0.932, h: 0.899, fill: W });
    tx(s, '“', { x: 3.938, y: 1.588, w: 1.2, h: 2.036, fontFace: PSB, fontSize: 115, color: VIO, wrap: false });
    tx(s, 'Music Doesn’t Lie. If There Is Something To Be Changed In This World, Then It Can Only Happen Through Music', {
        x: 3.792, y: 2.846, w: 7.485, h: 1.85, fontFace: LATOB, fontSize: 24, italic: true, lineSpacingMultiple: 1.5,
    });
    box(s, { x: 9.686, y: 5.348, w: 2.323, h: 0.572, fill: W });
    tx(s, 'QOUTE', { x: 9.97, y: 5.414, w: 1.755, h: 0.438, align: 'center', fontFace: PSB, fontSize: 20, color: DEEP, charSpacing: 0.5 });
});

/* 30 - THANK YOU */
slides.push(function (s) {
    s.background = { color: BG };
    accentPanel(s, 9.81, 0, 3.524, 7.5, 'd');
    box(s, { x: 7.057, y: 1.486, w: 5.41, h: 4.571, fill: PANEL, transparency: 42, rotate: 270 });
    s.addShape('line', { x: 1.272, y: 2.11, w: 3.912, h: 0, line: { color: W85, width: 3 } });
    s.addShape('line', { x: 2.81, y: 5.39, w: 3.912, h: 0, line: { color: W85, width: 3 } });
    tx(s, 'THANK', { x: 1.035, y: 2.215, w: 4.488, h: 1.313, fontFace: MBL, fontSize: 72, italic: true, color: W95 });
    tx(s, 'YOU', { x: 2.796, y: 3.674, w: 4.005, h: 1.717, fontFace: MBL, fontSize: 96, italic: true, color: W95 });
    iconTreble(s, 1.58, 4.454, 0.518, 1.407, MAG);
    iconNote(s, 5.612, 1.825, 0.574, 0.969, MAG);
});

/* ================================================================== assemble */
const pptx = new pptxgen();
pptx.defineLayout({ name: 'SPOTLIGHT16x9', width: 13.333, height: 7.5 });
pptx.layout = 'SPOTLIGHT16x9';
pptx.author = 'Spotlight';
pptx.title = 'Spotlight Event Organizer';

slides.forEach(function (build) { build(pptx.addSlide()); });

pptx.writeFile({ fileName: path.join(__dirname, '163e4058-0c54-47ed-a6ee-00d4566e1da7_grok_final.pptx') })
    .then(function (f) { console.log('wrote ' + f); });
