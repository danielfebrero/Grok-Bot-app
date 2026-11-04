/**
 * Sales Promotion deck — rebuilt with pptxgenjs.
 *
 * Slide size 13.333 x 7.5 in (16:9). All geometry is in inches, taken from the
 * source deck. Raster photos in the original are replaced by labelled
 * placeholder rectangles; the vector icons are redrawn with native shapes.
 */

const path = require('path');
const pptxgen = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const DARK = '143730'; // accent1 - deep green
const YELLOW = 'F0FD65'; // accent2 - lime
const LIGHT = 'E7E5E6'; // accent3 - warm grey
const WHITE = 'FFFFFF';
const BLACK = '000000';
const BAR_TEAL = '225C51'; // upper bar segment on the commercial-model chart

// Both slide backgrounds are perfectly linear (bilinear) ramps.
const BG_LIGHT = { tl: 'E7E5E6', tr: 'D9D6D7', bl: 'DFDCDE' };
const BG_DARK = { tl: '318776', tr: '1B4A41', bl: '246558' };
// Gradient used by the statistic cards.
const CARD_DARK = { tl: '20594E', tr: '318977', bl: '0E2722' };

const TITLE_FONT = 'Inter Medium';
const BODY_FONT = 'Public Sans';
const LABEL_FONT = 'Public Sans Medium';
const CONTACT_FONT = 'Funnel Display';

const LOREM = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.';
const FIGURES = 'Presentation Figures';

const SLIDE_W = 12192000 / 914400; // 13.3333in — same EMU size as the source deck
const SLIDE_H = 7.5;

// ---------------------------------------------------------- colour helpers ---
const toRgb = (hex) => [0, 2, 4].map((i) => parseInt(hex.substr(i, 2), 16));
const toHex = (rgb) =>
    rgb.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0').toUpperCase()).join('');

/** Colour of a linear ramp at (u,v), u/v in 0..1 across the ramp's box. */
function rampColor(ramp, u, v) {
    const tl = toRgb(ramp.tl);
    const tr = toRgb(ramp.tr);
    const bl = toRgb(ramp.bl);
    return toHex([0, 1, 2].map((i) => tl[i] + (tr[i] - tl[i]) * u + (bl[i] - tl[i]) * v));
}

/**
 * Paint a linear gradient over the whole slide using narrow strips that are
 * rotated perpendicular to the gradient direction (pptxgenjs has no gradient
 * fill, so the ramp is rasterised into shapes).
 */
function paintBackground(slide, ramp, bands = 80) {
    const tl = toRgb(ramp.tl);
    const tr = toRgb(ramp.tr);
    const bl = toRgb(ramp.bl);
    const gx = (tr[1] - tl[1]) / SLIDE_W; // green channel is the steepest, use it for direction
    const gy = (bl[1] - tl[1]) / SLIDE_H;
    const len = Math.hypot(gx, gy);
    const ux = gx / len;
    const uy = gy / len;
    const angle = (Math.atan2(uy, ux) * 180) / Math.PI;
    const proj = [[0, 0], [SLIDE_W, 0], [0, SLIDE_H], [SLIDE_W, SLIDE_H]].map(([x, y]) => x * ux + y * uy);
    const t0 = Math.min(...proj);
    const t1 = Math.max(...proj);
    const step = (t1 - t0) / bands;
    const mid = (SLIDE_W / 2) * ux + (SLIDE_H / 2) * uy;
    const strip = Math.hypot(SLIDE_W, SLIDE_H) + 1;

    for (let i = 0; i < bands; i++) {
        const t = t0 + step * (i + 0.5);
        const cx = SLIDE_W / 2 + (t - mid) * ux;
        const cy = SLIDE_H / 2 + (t - mid) * uy;
        slide.addShape('rect', {
            x: cx - (step + 0.02) / 2,
            y: cy - strip / 2,
            w: step + 0.02,
            h: strip,
            rotate: angle,
            fill: { color: rampColor(ramp, cx / SLIDE_W, cy / SLIDE_H) },
            line: { type: 'none' },
        });
    }
}

/**
 * Rounded card filled with a linear ramp. The card is sliced into horizontal
 * rows of cells; each row is inset so the outline follows the corner radius.
 * Rows are packed tightly inside the corner zones and coarser in the middle.
 */
function gradientCard(slide, x, y, w, h, ramp, cols = 14) {
    const r = 0.09836 * Math.min(w, h); // roundRect adj value used by the deck
    const edges = [0];
    while (edges[edges.length - 1] < h - 1e-6) {
        const d = Math.min(edges[edges.length - 1], h - edges[edges.length - 1]);
        edges.push(Math.min(h, edges[edges.length - 1] + (d < r ? r / 14 : 0.11)));
    }
    for (let i = 0; i < edges.length - 1; i++) {
        const y0 = edges[i];
        const rh = edges[i + 1] - y0;
        const edge = Math.min(y0 + rh / 2, h - y0 - rh / 2);
        const inset = edge >= r ? 0 : r - Math.sqrt(r * r - (r - edge) * (r - edge));
        const cw = (w - 2 * inset) / cols;
        for (let j = 0; j < cols; j++) {
            slide.addShape('rect', {
                x: x + inset + cw * j,
                y: y + y0,
                w: cw + (j === cols - 1 ? 0 : 0.01),
                h: rh + (i === edges.length - 2 ? 0 : 0.01),
                fill: { color: rampColor(ramp, (inset + cw * (j + 0.5)) / w, (y0 + rh / 2) / h) },
                line: { type: 'none' },
            });
        }
    }
}

// ------------------------------------------------------------ text helpers ---
const T = {
    title: { fontFace: TITLE_FONT, fontSize: 46, lineSpacingMultiple: 1.05 },
    body: { fontFace: BODY_FONT, fontSize: 14, lineSpacingMultiple: 1.5 },
    note: { fontFace: BODY_FONT, fontSize: 10.5, lineSpacingMultiple: 1.5 },
    stat: { fontFace: TITLE_FONT, fontSize: 44, lineSpacingMultiple: 1.5 },
    statSm: { fontFace: TITLE_FONT, fontSize: 32, lineSpacingMultiple: 1.5 },
    label: { fontFace: LABEL_FONT, fontSize: 11 },
};

/** Text box matching the source: top anchored, auto-height, no shape fill. */
function text(slide, str, opts) {
    slide.addText(str, Object.assign({ valign: 'top', fit: 'resize', color: BLACK }, opts));
}

const title = (slide, str, o) => text(slide, str, Object.assign({}, T.title, o));
const body = (slide, str, o) => text(slide, str, Object.assign({}, T.body, o));
const note = (slide, o) => text(slide, LOREM, Object.assign({}, T.note, o));
const figures = (slide, o) => text(slide, FIGURES, Object.assign({}, T.note, { h: 0.346, w: 1.762 }, o));

// ----------------------------------------------------------- icon helpers ----
/** Arrow glyph (0.238in artwork box); rotate 135 gives the down-left arrow. */
function arrow(slide, x, y, color, rotate = 0, size = 0.238) {
    slide.addShape('line', {
        x: x + size * 0.06,
        y: y + size / 2,
        w: size * 0.88,
        h: 0,
        rotate,
        line: { color, width: 1.25, endArrowType: 'triangle' },
    });
}

/** Circular badge with an arrow inside, as used all over the deck. */
function arrowBadge(slide, x, y, circleColor, arrowColor, rotate = 0) {
    slide.addShape('ellipse', { x, y, w: 0.406, h: 0.406, fill: { color: circleColor }, line: { type: 'none' } });
    arrow(slide, x + 0.084, y + 0.084, arrowColor, rotate);
}

/**
 * Calculator pictogram — the source deck uses a family of outlined
 * receipt/calculator/tag icons; this single glyph stands in for all of them.
 */
function docIcon(slide, x, y, color, size = 0.354) {
    const ln = { color, width: 1.1 };
    slide.addShape('roundRect', {
        x: x + size * 0.17, y: y + size * 0.05, w: size * 0.66, h: size * 0.9,
        rectRadius: size * 0.11, fill: { type: 'none' }, line: ln,
    });
    slide.addShape('rect', {
        x: x + size * 0.28, y: y + size * 0.16, w: size * 0.44, h: size * 0.16,
        fill: { color }, line: { type: 'none' },
    });
    for (let row = 0; row < 3; row++) {
        for (let col = 0; col < 3; col++) {
            slide.addShape('rect', {
                x: x + size * (0.29 + col * 0.16), y: y + size * (0.43 + row * 0.16),
                w: size * 0.09, h: size * 0.09, fill: { color }, line: { type: 'none' },
            });
        }
    }
}

/** Contact pictograms for the closing slide. */
function contactIcon(slide, kind, x, y, color, s = 0.315) {
    const ln = { color, width: 1.4 };
    const none = { type: 'none' };
    if (kind === 'mail') {
        slide.addShape('roundRect', {
            x, y: y + s * 0.17, w: s, h: s * 0.66, rectRadius: s * 0.08, fill: none, line: ln,
        });
        slide.addShape('line', { x: x + s * 0.06, y: y + s * 0.24, w: s * 0.44, h: s * 0.34, line: ln });
        slide.addShape('line', { x: x + s * 0.5, y: y + s * 0.24, w: s * 0.44, h: s * 0.34, flipH: true, line: ln });
    } else if (kind === 'phone') {
        // handset: rounded body with the ear/mouth pieces suggested by short bars
        slide.addShape('roundRect', {
            x: x + s * 0.2, y: y + s * 0.04, w: s * 0.6, h: s * 0.92, rectRadius: s * 0.16, fill: none, line: ln,
        });
        slide.addShape('line', { x: x + s * 0.38, y: y + s * 0.18, w: s * 0.24, h: 0, line: ln });
        slide.addShape('line', { x: x + s * 0.42, y: y + s * 0.8, w: s * 0.16, h: 0, line: ln });
    } else if (kind === 'pin') {
        slide.addShape('ellipse', { x: x + s * 0.14, y, w: s * 0.72, h: s * 0.72, fill: none, line: ln });
        slide.addShape('ellipse', {
            x: x + s * 0.38, y: y + s * 0.24, w: s * 0.24, h: s * 0.24, fill: none, line: ln,
        });
        slide.addShape('line', { x: x + s * 0.22, y: y + s * 0.58, w: s * 0.28, h: s * 0.4, line: ln });
        slide.addShape('line', { x: x + s * 0.5, y: y + s * 0.58, w: s * 0.28, h: s * 0.4, flipH: true, line: ln });
    } else {
        slide.addShape('ellipse', { x, y, w: s, h: s, fill: none, line: ln });
        slide.addShape('ellipse', { x: x + s * 0.31, y, w: s * 0.38, h: s, fill: none, line: ln });
        slide.addShape('line', { x, y: y + s * 0.5, w: s, h: 0, line: ln });
        slide.addShape('line', { x: x + s * 0.07, y: y + s * 0.25, w: s * 0.86, h: 0, line: ln });
        slide.addShape('line', { x: x + s * 0.07, y: y + s * 0.75, w: s * 0.86, h: 0, line: ln });
    }
}

/**
 * Stand-in for a photo in the source deck (both remaining photos are device
 * mock-ups): a dark rounded bezel around a light screen, plus an "[image]"
 * caption so the substitution is obvious.
 */
function imagePlaceholder(slide, x, y, w, h, radius, bezel) {
    slide.addShape('roundRect', {
        x, y, w, h, rectRadius: radius,
        fill: { color: '111111' }, line: { type: 'none' },
    });
    slide.addShape('roundRect', {
        x: x + bezel, y: y + bezel, w: w - 2 * bezel, h: h - 2 * bezel, rectRadius: radius * 0.8,
        fill: { color: 'D7D4D5' }, line: { type: 'none' },
    });
    const cy = Math.min(y + h / 2, SLIDE_H - 0.9); // keep the caption on-slide
    slide.addText('[image]', {
        x, y: cy - 0.25, w, h: 0.5,
        align: 'center', valign: 'middle', fontFace: BODY_FONT, fontSize: 12, color: '8A8A8A',
    });
}

// ------------------------------------------------------------ slide chrome ---
/**
 * Header strip shared by every slide: page counter, section label, the
 * "SALES PROMOTION" pill and the round arrow badge in the corner.
 */
function header(slide, o) {
    const left = o.leftColor || BLACK;
    const pill = o.pillColor || BLACK;
    text(slide, '00/00', Object.assign({ x: 0.734, y: 0.538, w: 0.644, h: 0.286, wrap: false, color: left }, T.label));
    text(slide, o.label, Object.assign({ x: 1.715, y: 0.538, w: o.labelW, h: 0.286, wrap: false, color: left }, T.label));
    text(slide, 'SALES PROMOTION',
        Object.assign({ x: 9.816, y: 0.538, w: 2.009, h: 0.286, align: 'center', color: pill }, T.label));
    slide.addShape('roundRect', {
        x: 9.816, y: 0.479, w: 2.009, h: 0.406, rectRadius: 0.203,
        fill: { type: 'none' }, line: { color: pill, width: 0.75 },
    });
    arrowBadge(slide, 12.089, 0.479, o.badge, o.badgeArrow, 135);
}

// ----------------------------------------------------------------- slides ----
const pres = new pptxgen();
pres.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
pres.layout = 'WIDE';
pres.author = 'pptxgenjs';
pres.title = 'Sales Promotion';

const newSlide = (ramp) => {
    const s = pres.addSlide();
    if (ramp) paintBackground(s, ramp);
    return s;
};
const flatSlide = (color) => {
    const s = pres.addSlide();
    s.background = { color };
    return s;
};

const LIGHT_HEAD = { leftColor: BLACK, pillColor: BLACK, badge: DARK, badgeArrow: WHITE };
const DARK_HEAD = { leftColor: WHITE, pillColor: WHITE, badge: YELLOW, badgeArrow: DARK };

// --- 1. cover -----------------------------------------------------------------
function slide01() {
    const s = newSlide(BG_LIGHT);
    header(s, Object.assign({ label: 'PROMOTION OVERVIEW', labelW: 1.994 }, LIGHT_HEAD));
    title(s, 'Driving Sales Through Strategic Promotions',
        { x: 0.704, y: 1.486, w: 6.12, h: 2.491, color: DARK });
    body(s, 'This strategy addresses visibility gaps, conversion inefficiencies, and retention challenges '
        + 'through a focused and data-informed growth marketing approach.',
        { x: 0.742, y: 4.49, w: 5.766, h: 1.118 });
    note(s, { x: 1.715, y: 6.174, w: 1.933, h: 0.864 });
    arrowBadge(s, 0.839, 6.329, DARK, WHITE);
}

// --- 2. intro with yellow stat card -------------------------------------------
function slide02() {
    const s = newSlide(BG_DARK);
    header(s, Object.assign({ label: 'PROMOTION OVERVIEW', labelW: 2.117 }, DARK_HEAD));
    s.addShape('roundRect', {
        x: 0.839, y: 3.75, w: 2.993, h: 3.229, rectRadius: 0.294,
        fill: { color: YELLOW }, line: { type: 'none' },
    });
    text(s, '45%', Object.assign({}, T.stat, { x: 1.024, y: 3.672, w: 1.762, h: 1.082, color: DARK }));
    figures(s, { x: 1.024, y: 4.7, color: DARK });
    s.addShape('roundRect', {
        x: 2.943, y: 4.01, w: 0.602, h: 0.587, rectRadius: 0.098,
        fill: { color: DARK }, line: { type: 'none' },
    });
    docIcon(s, 3.067, 4.126, YELLOW);
    note(s, { x: 1.024, y: 5.825, w: 1.933, h: 0.864, color: DARK });

    title(s, 'Setting the Focus for Sales Acceleration',
        { x: 0.704, y: 1.486, w: 6.704, h: 1.678, color: WHITE });
    body(s, 'This presentation introduces a promotion-focused framework aimed at aligning offers, timing, '
        + 'and channels to support revenue growth and market responsiveness.',
        { x: 8.748, y: 3.607, w: 3.747, h: 1.825, color: WHITE });
    note(s, { x: 9.708, y: 1.522, w: 1.933, h: 0.864, color: WHITE });
    arrowBadge(s, 8.832, 1.663, YELLOW, DARK);
}

// --- 3. agenda pills ----------------------------------------------------------
function slide03() {
    const s = newSlide(BG_LIGHT);
    header(s, Object.assign({ label: 'PROMOTION OVERVIEW', labelW: 1.994 }, LIGHT_HEAD));
    title(s, 'What This Promotion Strategy Covers',
        { x: 2.847, y: 1.486, w: 7.639, h: 1.678, align: 'center', color: DARK });

    const items = [
        ['1. Promotion Overview', 0.839, 3.977, DARK],
        ['2. Market Context', 4.835, 3.977, YELLOW],
        ['3. Strategic Direction', 8.832, 3.977, DARK],
        ['4. Promotion Execution', 0.839, 5.093, YELLOW],
        ['5. Team & Resources', 4.835, 5.093, DARK],
        ['6. Commercial Model', 8.832, 5.093, YELLOW],
        ['7. Next Steps', 4.835, 6.139, YELLOW],
    ];
    items.forEach(([label, x, y, fill]) => {
        s.addShape('roundRect', {
            x, y, w: 3.663, h: 0.622, rectRadius: 0.311,
            fill: { color: fill }, line: { type: 'none' },
        });
        text(s, label, {
            x, y: y - 0.022, w: 3.663, h: 0.552, align: 'center',
            fontFace: BODY_FONT, fontSize: 20, lineSpacingMultiple: 1.5,
            color: fill === DARK ? WHITE : DARK,
        });
    });
}

// --- 4. market context --------------------------------------------------------
function slide04() {
    const s = newSlide(BG_LIGHT);
    header(s, { label: 'MARKET CONTEXT', labelW: 1.617, leftColor: BLACK, pillColor: WHITE, badge: DARK, badgeArrow: WHITE });
    title(s, 'Understanding the Sales Environment', { x: 0.704, y: 1.486, w: 6.12, h: 1.678 });
    body(s, 'We analyze market conditions, customer behavior, and competitive dynamics to ensure sales '
        + 'promotions are relevant, timely, and aligned with business goals.',
        { x: 0.781, y: 5.114, w: 4.435, h: 1.471 });
    text(s, '75%', Object.assign({}, T.stat, { x: 1.674, y: 3.778, w: 1.762, h: 1.082 }));
    s.addShape('roundRect', {
        x: 0.849, y: 4.141, w: 0.602, h: 0.587, rectRadius: 0.098,
        fill: { color: DARK }, line: { type: 'none' },
    });
    docIcon(s, 0.973, 4.257, WHITE);
    figures(s, { x: 3.385, y: 4.245 });
}

// --- 5. barriers, two dark stat cards -----------------------------------------
function slide05() {
    const s = flatSlide(YELLOW);
    header(s, Object.assign({ label: 'MARKET CONTEXT', labelW: 1.617 }, LIGHT_HEAD));
    title(s, 'Identifying Barriers to Sales Growth',
        { x: 0.704, y: 5.382, w: 6.12, h: 1.678, color: DARK });
    body(s, 'This strategy addresses demand fluctuations, conversion challenges, and customer hesitation '
        + 'through targeted promotional tactics and clear value communication.',
        { x: 0.734, y: 1.507, w: 4.27, h: 1.825 });

    [[7.549, 1.663, '54%'], [10.228, 4.551, '82%']].forEach(([x, y, value]) => {
        gradientCard(s, x, y, 2.267, 2.441, CARD_DARK);
        s.addShape('roundRect', {
            x: x + 0.265, y: y + 0.235, w: 0.516, h: 0.518, rectRadius: 0.086,
            fill: { color: YELLOW }, line: { type: 'none' },
        });
        docIcon(s, x + 0.346, y + 0.317, DARK);
        text(s, value, Object.assign({}, T.stat, { x: x + 0.166, y: y + 0.853, w: 1.762, h: 1.082, color: WHITE }));
        figures(s, { x: x + 0.177, y: y + 1.881, color: WHITE });
    });

    note(s, { x: 1.714, y: 4.447, w: 2.267, h: 0.599 });
    arrowBadge(s, 0.839, 4.57, DARK, WHITE);
}

// --- 6. vision / mission ------------------------------------------------------
function slide06() {
    const s = newSlide(BG_LIGHT);
    header(s, Object.assign({ label: 'MARKET CONTEXT', labelW: 1.617 }, LIGHT_HEAD));
    title(s, 'Aligning Promotions with Business Direction',
        { x: 6.667, y: 1.486, w: 6.12, h: 2.491, color: DARK });
    body(s, 'Promotional activities are aligned with long-term business vision and mission to ensure '
        + 'short-term sales growth supports sustainable brand value.',
        { x: 6.705, y: 4.458, w: 5.766, h: 1.118 });
    [[6.705, '94%'], [8.714, '120+']].forEach(([x, value]) => {
        text(s, value, Object.assign({}, T.statSm, { x, y: 5.943, w: 1.414, h: 0.815, color: DARK }));
        figures(s, { x, y: 6.704 });
    });
}

// --- 7. objectives timeline ---------------------------------------------------
function slide07() {
    const s = flatSlide(YELLOW);
    header(s, Object.assign({ label: 'STRATEGIC DIRECTION', labelW: 1.939 }, LIGHT_HEAD));
    title(s, 'Clear Goals Behind Each Promotion',
        { x: 3.606, y: 1.486, w: 6.12, h: 1.678, align: 'center', color: DARK });
    body(s, 'Promotion objectives focus on increasing sales volume, driving urgency, attracting new '
        + 'customers, and supporting retention through measurable targets.',
        { x: 2.748, y: 6.257, w: 7.837, h: 0.765, align: 'center' });

    s.addShape('line', { x: 0, y: 4.157, w: SLIDE_W, h: 0, line: { color: DARK, width: 1 } });
    [1.486, 4.049, 6.612, 9.175, 11.737].forEach((x) => {
        s.addShape('ellipse', { x, y: 4.102, w: 0.11, h: 0.11, fill: { color: DARK }, line: { type: 'none' } });
    });
    [0.66, 3.223, 5.786, 8.348, 10.911].forEach((x) => {
        text(s, '30%', Object.assign({}, T.statSm, { x, y: 4.376, w: 1.762, h: 0.815, align: 'center' }));
        figures(s, { x, y: 5.137, align: 'center' });
    });
}

// --- 8. approach (dark) -------------------------------------------------------
function slide08() {
    const s = newSlide(BG_DARK);
    header(s, Object.assign({ label: 'STRATEGIC DIRECTION', labelW: 1.939 }, DARK_HEAD));
    title(s, 'Our Approach to Sales Stimulation',
        { x: 0.701, y: 1.494, w: 6.12, h: 1.678, color: WHITE });
    s.addShape('roundRect', {
        x: 0.847, y: 3.485, w: 2.009, h: 0.406, rectRadius: 0.203,
        fill: { color: YELLOW }, line: { type: 'none' },
    });
    text(s, 'LET\u2019S JOIN US!',
        Object.assign({ x: 0.847, y: 3.545, w: 2.009, h: 0.286, align: 'center', color: DARK }, T.label));
    body(s, 'The strategy combines audience segmentation, offer design, channel selection, and timing '
        + 'optimization to maximize promotional effectiveness.',
        { x: 0.742, y: 4.443, w: 4.744, h: 1.471, color: WHITE });
}

// --- 9. phases with fading keyword stack --------------------------------------
function slide09() {
    const s = newSlide(BG_LIGHT);
    header(s, Object.assign({ label: 'STRATEGIC DIRECTION', labelW: 1.939 }, LIGHT_HEAD));
    title(s, 'Key Phases of Promotional Activities', { x: 0.704, y: 1.486, w: 7.143, h: 1.678, color: DARK });
    body(s, 'Promotion milestones include planning, offer development, launch execution, monitoring, '
        + 'optimization, and post-promotion evaluation stages.',
        { x: 1.715, y: 5.93, w: 5.46, h: 1.118 });

    // y, height, font size, opacity of the repeated "Keyword Option" label
    const keywords = [
        [1.663, 0.3, 12, 80], [2.389, 0.367, 16, 60], [3.181, 0.433, 20, 35],
        [4.039, 0.566, 28, 0],
        [5.03, 0.433, 20, 35], [5.888, 0.367, 16, 60], [6.68, 0.3, 12, 80],
    ];
    keywords.forEach(([y, h, size, transparency]) => {
        text(s, 'Keyword Option', {
            x: 8.477, y, w: 4.017, h, align: 'center',
            fontFace: TITLE_FONT, fontSize: size, lineSpacingMultiple: 1.05,
            color: DARK, transparency,
        });
    });

    [[1.715, 4.115, 0.839, 4.269], [4.734, 4.115, 3.857, 4.269]].forEach(([tx, ty, bx, by]) => {
        note(s, { x: tx, y: ty, w: 1.933, h: 0.864 });
        arrowBadge(s, bx, by, DARK, WHITE);
    });
    arrowBadge(s, 0.839, 6.052, DARK, WHITE);
}

// --- 10. services, yellow + dark cards ----------------------------------------
function slide10() {
    const s = newSlide(BG_LIGHT);
    header(s, Object.assign({ label: 'PROMOTION EXECUTION', labelW: 2.064 }, LIGHT_HEAD));
    title(s, 'Services Supporting Sales Promotions', { x: 6.667, y: 1.486, w: 6.12, h: 2.491, color: DARK });
    body(s, 'Services include promotion planning, creative asset development, channel activation, '
        + 'performance monitoring, and optimization throughout promotion cycles.',
        { x: 0.742, y: 1.486, w: 5.766, h: 1.118 });

    gradientCard(s, 9.562, 4.538, 2.267, 2.441, CARD_DARK);
    s.addShape('roundRect', {
        x: 9.827, y: 4.773, w: 0.516, h: 0.518, rectRadius: 0.086, fill: { color: YELLOW }, line: { type: 'none' },
    });
    docIcon(s, 9.908, 4.855, DARK);
    text(s, '324', Object.assign({}, T.stat, { x: 9.728, y: 5.391, w: 1.762, h: 1.082, color: WHITE }));
    figures(s, { x: 9.739, y: 6.419, color: WHITE });

    s.addShape('roundRect', {
        x: 4.118, y: 4.563, w: 2.267, h: 2.441, rectRadius: 0.223, fill: { color: YELLOW }, line: { type: 'none' },
    });
    s.addShape('roundRect', {
        x: 4.372, y: 4.773, w: 0.516, h: 0.518, rectRadius: 0.086, fill: { color: DARK }, line: { type: 'none' },
    });
    docIcon(s, 4.453, 4.855, WHITE);
    text(s, '64%', Object.assign({}, T.stat, { x: 4.273, y: 5.391, w: 1.762, h: 1.082, color: DARK }));
    figures(s, { x: 4.284, y: 6.419, color: DARK });

    [[4.428, 4.582], [6.174, 6.329]].forEach(([ty, by]) => {
        note(s, { x: 1.715, y: ty, w: 1.933, h: 0.864 });
        arrowBadge(s, 0.839, by, DARK, WHITE);
    });
}

// --- 11. implementation (yellow) ----------------------------------------------
function slide11() {
    const s = flatSlide(YELLOW);
    header(s, Object.assign({ label: 'PROMOTION EXECUTION', labelW: 2.064 }, LIGHT_HEAD));
    title(s, 'How Promotions Are Implemented',
        { x: 3.606, y: 1.486, w: 6.12, h: 1.678, align: 'center', color: DARK });
    body(s, 'Execution follows structured workflows with defined roles, timelines, and checkpoints to '
        + 'ensure promotions run efficiently and consistently.',
        { x: 3.784, y: 3.541, w: 5.766, h: 1.118, align: 'center' });
    note(s, { x: 0.734, y: 1.543, w: 1.933, h: 0.864 });
    note(s, { x: 10.667, y: 1.543, w: 1.933, h: 0.864, align: 'right' });
}

// --- 12. tools and materials ---------------------------------------------------
function slide12() {
    const s = newSlide(BG_LIGHT);
    header(s, Object.assign({ label: 'PROMOTION EXECUTION', labelW: 2.064 }, LIGHT_HEAD));
    title(s, 'Tools and Materials Driving Sales', { x: 0.704, y: 1.486, w: 6.676, h: 1.678, color: DARK });
    body(s, 'Promotion assets include visual materials, messaging frameworks, digital content, and '
        + 'support tools designed to communicate offers clearly.',
        { x: 0.734, y: 5.944, w: 5.766, h: 1.118 });
    [[0.704, '47%', 1.762], [3.379, '540+', 2.214]].forEach(([x, value, w]) => {
        text(s, value, Object.assign({}, T.stat, { x, y: 3.75, w, h: 1.082, color: DARK }));
        figures(s, { x, y: 4.778 });
    });
}

// --- 13. why it works (dark) ---------------------------------------------------
function slide13() {
    const s = newSlide(BG_DARK);
    header(s, Object.assign({ label: 'PROMOTION EXECUTION', labelW: 2.064 }, DARK_HEAD));
    title(s, 'Why This Promotion Strategy Works', { x: 0.704, y: 1.486, w: 6.513, h: 1.678, color: WHITE });
    body(s, 'This approach delivers value through focused offers, clear communication, efficient '
        + 'execution, and measurable sales performance.',
        { x: 0.734, y: 3.835, w: 3.617, h: 1.471, color: WHITE });
    [[1.715, 0.839], [4.708, 3.832]].forEach(([tx, bx]) => {
        note(s, { x: tx, y: 6.174, w: 1.933, h: 0.864, color: WHITE });
        arrowBadge(s, bx, 6.329, YELLOW, DARK);
    });
}

// --- 14. team -----------------------------------------------------------------
function slide14() {
    const s = newSlide(BG_LIGHT);
    header(s, Object.assign({ label: 'TEAM & RESOURCES', labelW: 1.766 }, LIGHT_HEAD));
    title(s, 'The Team Managing Promotions',
        { x: 0.839, y: 1.486, w: 11.656, h: 0.865, align: 'center', color: DARK });
    body(s, 'A cross-functional team of marketers, sales coordinators, and creatives collaborates to '
        + 'plan, execute, and evaluate promotional initiatives.',
        { x: 2.889, y: 2.696, w: 7.727, h: 0.765, align: 'center' });
    [0.839, 3.945, 7.052, 10.158].forEach((x) => {
        body(s, 'Owner Name', { x, y: 6.661, w: 2.336, h: 0.411, align: 'center' });
    });
}

// --- 15. collaboration ---------------------------------------------------------
function slide15() {
    const s = newSlide(BG_LIGHT);
    header(s, { label: 'TEAM & RESOURCES', labelW: 1.766, leftColor: WHITE, pillColor: BLACK, badge: DARK, badgeArrow: WHITE });
    title(s, 'How Teams Coordinate Promotion Efforts',
        { x: 6.69, y: 3.7, w: 6.12, h: 2.491, color: DARK });
    body(s, 'Collaboration is built on clear communication, aligned timelines, shared targets, and '
        + 'regular performance updates throughout promotion periods.',
        { x: 6.729, y: 1.501, w: 5.766, h: 1.118 });
    s.addShape('roundRect', {
        x: 6.843, y: 6.567, w: 2.009, h: 0.406, rectRadius: 0.203, fill: { color: DARK }, line: { type: 'none' },
    });
    text(s, 'LET\u2019S JOIN US!',
        Object.assign({ x: 6.843, y: 6.627, w: 2.009, h: 0.286, align: 'center', color: WHITE }, T.label));
}

// --- 16. pricing bar chart -----------------------------------------------------
function slide16() {
    const s = flatSlide(YELLOW);
    header(s, Object.assign({ label: 'COMMERCIAL MODEL', labelW: 1.838 }, LIGHT_HEAD));
    title(s, 'Structuring Promotional Value', { x: 0.704, y: 1.486, w: 4.152, h: 2.491, color: DARK });
    body(s, 'Pricing and promotional offers are structured based on margins, target segments, campaign '
        + 'duration, and desired sales outcomes.',
        { x: 5.739, y: 1.559, w: 3.891, h: 1.471 });
    note(s, { x: 1.715, y: 5.607, w: 1.762, h: 0.864 });
    note(s, { x: 3.724, y: 5.607, w: 1.762, h: 0.864 });
    arrowBadge(s, 0.839, 5.762, DARK, WHITE);

    // x, width, top of the teal column, top of the dark base, label text, label y
    const bars = [
        [6.193, 1.318, 5.5, 7.178, '14%', 5.713],
        [7.847, 1.318, 4.094, 5.941, '22%', 4.346],
        [9.502, 1.339, 2.979, 4.826, '54%', 3.242],
        [11.156, 1.339, 1.663, 3.51, '72%', 1.923],
    ];
    bars.forEach(([x, w, top]) => {
        s.addShape('rect', { x, y: top, w, h: SLIDE_H - top, fill: { color: BAR_TEAL }, line: { type: 'none' } });
    });
    bars.forEach(([x, w, , , label, labelY]) => {
        text(s, label, {
            x, y: labelY, w, h: 0.586, align: 'center',
            fontFace: TITLE_FONT, fontSize: 32, lineSpacingMultiple: 0.9, color: WHITE,
        });
    });
    bars.forEach(([x, w, , base]) => {
        s.addShape('rect', { x, y: base, w, h: SLIDE_H - base, fill: { color: DARK }, line: { type: 'none' } });
    });
}

// --- 17. return on promotions ---------------------------------------------------
function slide17() {
    const s = newSlide(BG_LIGHT);
    header(s, Object.assign({ label: 'COMMERCIAL MODEL', labelW: 1.838 }, LIGHT_HEAD));
    title(s, 'Maximizing Return from Promotions',
        { x: 3.606, y: 1.467, w: 6.12, h: 1.678, align: 'center', color: DARK });
    body(s, 'PLACEHOLDER'
        + 'value and supporting long-term customer relationships.',
        { x: 2.738, y: 3.591, w: 7.857, h: 0.765, align: 'center' });
}

// --- 18. wrap-up with tablet mock-up --------------------------------------------
function slide18() {
    const s = newSlide(BG_LIGHT);
    // Source has a rotated tablet photo running off the right/bottom edge.
    imagePlaceholder(s, 6.733, 3.207, 7.0, 4.6, 0.32, 0.055);
    header(s, Object.assign({ label: 'NEXT STEPS', labelW: 1.147 }, LIGHT_HEAD));
    title(s, 'Bringing Planning and Execution Together',
        { x: 0.704, y: 1.486, w: 6.12, h: 2.491, color: DARK });
    body(s, 'This sales promotion strategy provides a clear framework to support planning, execution, '
        + 'and evaluation of promotional initiatives.',
        { x: 0.742, y: 4.638, w: 5.235, h: 1.118 });
    note(s, { x: 8.724, y: 1.507, w: 1.762, h: 0.864 });
    note(s, { x: 10.732, y: 1.507, w: 1.762, h: 0.864 });
    arrowBadge(s, 7.847, 1.661, DARK, WHITE);
}

// --- 19. contact ----------------------------------------------------------------
function slide19() {
    const s = newSlide(BG_LIGHT);
    // Source has a phone mock-up photo running off the bottom edge.
    imagePlaceholder(s, 0.95, 1.62, 4.1, 6.5, 0.6, 0.06);
    header(s, Object.assign({ label: 'NEXT STEPS', labelW: 1.147 }, LIGHT_HEAD));
    title(s, 'Let\u2019s Discuss Your Promotion Plan', { x: 6.374, y: 1.486, w: 6.12, h: 1.678, color: DARK });

    const rows = [
        ['Email', ':       yourmail@email.com', 'mail'],
        ['Phone', ':      +123456789', 'phone'],
        ['Address', ':      123 Your Street, Your City', 'pin'],
        ['Website', ':      www.yourwebsite.com', 'globe'],
    ];
    const para = (str) => ({ text: str, options: { lineSpacingMultiple: 2.5, breakLine: true } });
    text(s, rows.map((r) => para(r[0])),
        { x: 7.25, y: 3.78, w: 2.926, h: 3.002, fontFace: CONTACT_FONT, fontSize: 18 });
    text(s, rows.map((r) => para(r[1])),
        { x: 8.442, y: 3.78, w: 3.76, h: 3.002, fontFace: CONTACT_FONT, fontSize: 18 });
    rows.forEach((r, i) => contactIcon(s, r[2], 6.532, 4.148 + i * 0.7517, DARK));
}

// --- 20. closing (dark) ----------------------------------------------------------
function slide20() {
    const s = newSlide(BG_DARK);
    header(s, Object.assign({ label: 'NEXT STEPS', labelW: 1.147 }, DARK_HEAD));
    title(s, 'Ready to Boost Sales Performance', { x: 0.704, y: 1.486, w: 6.12, h: 1.678, color: WHITE });
    body(s, 'We look forward to supporting your sales growth through structured promotion strategies, '
        + 'efficient execution, and performance-driven decision-making.',
        { x: 0.742, y: 3.777, w: 5.766, h: 1.118, color: WHITE });
    note(s, { x: 1.768, y: 5.582, w: 1.762, h: 0.864, color: WHITE });
    arrowBadge(s, 0.839, 5.737, YELLOW, DARK);

    [[7.549, YELLOW, '26%'], [10.221, LIGHT, '73%']].forEach(([x, fill, value]) => {
        s.addShape('roundRect', {
            x, y: 4.538, w: 2.267, h: 2.441, rectRadius: 0.223, fill: { color: fill }, line: { type: 'none' },
        });
        s.addShape('roundRect', {
            x: x + 0.265, y: 4.773, w: 0.516, h: 0.518, rectRadius: 0.086,
            fill: { color: DARK }, line: { type: 'none' },
        });
        docIcon(s, x + 0.346, 4.855, WHITE);
        text(s, value, Object.assign({}, T.stat, { x: x + 0.166, y: 5.391, w: 1.762, h: 1.082, color: DARK }));
        figures(s, { x: x + 0.177, y: 6.419, color: DARK });
    });
}

[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((build) => build());

pres.writeFile({ fileName: path.join(__dirname, '12f2e1a6-c020-4726-b9a3-abaab85f9426_grok_final.pptx') })
    .then((f) => console.log('wrote', f));
