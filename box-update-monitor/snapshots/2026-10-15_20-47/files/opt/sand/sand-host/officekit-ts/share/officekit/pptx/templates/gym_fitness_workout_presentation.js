/**
 * WORKOUT — Gym & Fitness Presentation Template (30 slides, 13.333" x 7.5")
 * Recreated with pptxgenjs. Photographic content of the original deck is
 * replaced by flat "[image]" placeholder rectangles.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Design tokens
 * ------------------------------------------------------------------ */
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const DARK = '1A1A1A'; // accent1 - deck background
const PANEL = '313131'; // accent1 lum 90/10 - raised panel
const LIME = 'C8F31D'; // accent2 - brand colour
const LIME_D = '9DC20A'; // accent2 lum 75
const LIME_DD = '698107'; // accent2 lum 50
const LIME_L = 'F4FDD2'; // accent2 lum 20/80
const WHITE = 'FFFFFF';
const INK = '404040'; // tx1 lum 75/25 - text on light shapes
const GRAY = 'A5A5A5'; // accent3
const SILVER = 'E8E8E8'; // accent1 lum 10/90

const HEAD = 'Kanit SemiBold'; // major latin font
const BODY = 'Roboto'; // minor latin font

const NO_LINE = { type: 'none' };
const CARD_SHADOW = { type: 'outer', color: '000000', opacity: 0.2, blur: 25, offset: 10, angle: 50 };
const SOFT_SHADOW = { type: 'outer', color: '000000', opacity: 0.34, blur: 72, offset: 30, angle: 90 };

/* ------------------------------------------------------------------ *
 * Low level helpers
 * ------------------------------------------------------------------ */

/** Linear blend between two hex colours (t = 0 -> a, 1 -> b). */
function mix(a, b, t) {
	t = Math.max(0, Math.min(1, t));
	let out = '';
	for (let i = 0; i < 3; i++) {
		const ca = parseInt(a.substr(i * 2, 2), 16);
		const cb = parseInt(b.substr(i * 2, 2), 16);
		out += Math.round(ca + (cb - ca) * t).toString(16).toUpperCase().padStart(2, '0');
	}
	return out;
}

/** Free-form polygon. `pts` are [x,y] pairs in inches, relative to x/y. */
function poly(slide, o) {
	const opts = {
		x: o.x, y: o.y, w: o.w, h: o.h,
		points: o.pts.map(p => ({ x: p[0], y: p[1] })).concat([{ close: true }]),
		fill: o.fill, line: o.line || NO_LINE,
	};
	if (o.flipH) opts.flipH = true;
	if (o.flipV) opts.flipV = true;
	if (o.rotate) opts.rotate = o.rotate;
	if (o.shadow) opts.shadow = o.shadow;
	if (o.text) {
		Object.assign(opts, { shape: 'custGeom', fontFace: BODY, fontSize: 12, color: INK, align: 'center', valign: 'middle' }, o.textOpts || {});
		slide.addText(o.text, opts);
	} else {
		slide.addShape('custGeom', opts);
	}
}

/* Local-space geometry builders (all return [x,y] pairs in inches). */
const geo = {
	// Slanted box; `adj` is the PowerPoint parallelogram adjust value (0..1).
	parallelogram: (w, h, adj) => {
		const o = Math.min(w, h) * adj;
		return [[0, h], [o, 0], [w, 0], [w - o, h]];
	},
	// Arrow: chevron.
	chevron: (w, h, adj) => {
		const o = Math.min(w, h) * adj;
		return [[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h], [o, h / 2]];
	},
	// Arrow: pentagon.
	homePlate: (w, h, adj) => {
		const o = Math.min(w, h) * adj;
		return [[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h]];
	},
	// Rectangle with two opposite corners snipped.
	snip2Diag: (w, h, a1, a2) => {
		const s = Math.min(w, h), d1 = s * a1, d2 = s * a2;
		return [[d1, 0], [w - d2, 0], [w, d2], [w, h - d1], [w - d1, h], [d2, h], [0, h - d2], [0, d1]];
	},
	rtTriangle: (w, h) => [[0, h], [0, 0], [w, h]],
	// Scale a list of fractional [0..1] coordinates into inches.
	frac: (w, h, pts) => pts.map(p => [p[0] * w, p[1] * h]),
};

/** Clip a polygon to the horizontal band y0..y1 (Sutherland-Hodgman). */
function clipBand(pts, y0, y1) {
	const clip = (poly, keep, coord) => {
		const out = [];
		for (let i = 0; i < poly.length; i++) {
			const a = poly[i], b = poly[(i + 1) % poly.length];
			const ka = keep(a), kb = keep(b);
			if (ka) out.push(a);
			if (ka !== kb) {
				const t = (coord - a[1]) / (b[1] - a[1]);
				out.push([a[0] + (b[0] - a[0]) * t, coord]);
			}
		}
		return out;
	};
	let p = clip(pts, q => q[1] >= y0, y0);
	if (p.length < 3) return [];
	return clip(p, q => q[1] <= y1, y1);
}

/**
 * The deck's signature "glass" ramp: lime at 92% of the axis fading to fully
 * transparent white at 8%, composited over the dark background. Returns the
 * flat colour that this ramp resolves to at position t (0..1).
 */
function glassRamp(t) {
	const a = Math.max(0, Math.min(1, (t - 0.08) / 0.84));
	return mix(mix(WHITE, LIME, a), DARK, 1 - a);
}

/**
 * Vertical gradient inside an arbitrary polygon, painted as horizontal bands.
 * (pptxgenjs has no gradient fill, so the ramp is drawn explicitly.)
 */
function gradPoly(slide, o) {
	const n = o.bands || 16, step = o.h / n, lap = step * 0.7; // overlap hides seams
	for (let i = 0; i < n; i++) {
		const band = clipBand(o.pts, step * i - lap, step * (i + 1) + lap);
		if (band.length < 3) continue;
		const t = (i + 0.5) / n; // 0 = top of shape
		poly(slide, {
			x: o.x, y: o.y, w: o.w, h: o.h, pts: band, flipH: o.flipH,
			fill: { color: glassRamp(o.up ? 1 - t : t) },
		});
	}
}

/** Rotated bar filled with a gradient running along its长 axis. */
function gradBar(slide, o) {
	const n = o.bands || 12;
	const th = ((o.rotate || 0) * Math.PI) / 180;
	const cx = o.x + o.w / 2, cy = o.y + o.h / 2, bw = o.w / n;
	for (let i = 0; i < n; i++) {
		const d = (i + 0.5) * bw - o.w / 2;
		const t = Math.max(0, ((i + 0.5) / n - (o.stop || 0)) / (0.8 - (o.stop || 0)));
		slide.addShape('rect', {
			x: cx + d * Math.cos(th) - bw / 2, y: cy + d * Math.sin(th) - o.h / 2,
			w: bw * 2.4, h: o.h, rotate: o.rotate,
			fill: { color: mix(o.from, o.to, t) }, line: NO_LINE,
		});
	}
}

/* ------------------------------------------------------------------ *
 * Deck furniture
 * ------------------------------------------------------------------ */

/* Outline of the dumbbell mark used for the logo and the round icons. */
const DUMBBELL = [
	[425, 0], [358, 0], [343, 14], [343, 44], [326, 62], [301, 61], [224, 47],
	[148, 61], [123, 62], [106, 44], [106, 14], [91, 0], [24, 0], [10, 14],
	[10, 44], [8, 53], [0, 74], [4, 98], [8, 120], [10, 131], [10, 160],
	[24, 173], [91, 173], [106, 160], [106, 130], [123, 112], [148, 113],
	[224, 126], [301, 113], [326, 112], [343, 130], [343, 160], [358, 173],
	[425, 173], [439, 160], [439, 143], [448, 98], [443, 72], [439, 43],
	[439, 14], [433, 0],
];

/** Two crossed dumbbells; `w` is the width of the whole mark (0.518 = logo). */
function dumbbells(slide, x, y, w, color) {
	const s = w / 0.518, bw = 0.314 * s, bh = 0.121 * s;
	const pts = DUMBBELL.map(p => [(p[0] / 448) * bw, (p[1] / 173) * bh]);
	poly(slide, { x: x + 0.204 * s, y: y + 0.109 * s, w: bw, h: bh, pts, fill: { color } });
	poly(slide, { x, y, w: bw, h: bh, pts, fill: { color }, rotate: 330.28 });
}

/** Three slanted speed stripes, top-right corner. */
function stripes(slide, color) {
	[[12.412, 0.434], [12.220, 0.379], [12.055, 0.351]].forEach(([x, w]) => {
		poly(slide, { x, y: 0.599, w, h: 0.225, pts: geo.parallelogram(w, 0.225, 0.99762), fill: { color } });
	});
}

/** Logo lock-up (top-left) plus optional speed stripes (top-right). */
function chrome(slide, opts) {
	const o = opts || {};
	slide.addText('WORKOUT', {
		x: 1.155, y: 0.503, w: 1.672, h: 0.303,
		fontFace: HEAD, fontSize: 12, color: WHITE, charSpacing: 3, valign: 'top',
	});
	dumbbells(slide, 0.499, 0.482, 0.518, LIME);
	if (o.stripes !== false) stripes(slide, o.stripeColor || LIME);
}

/* ------------------------------------------------------------------ *
 * Text helpers
 * ------------------------------------------------------------------ */

/** Small lime kicker above a headline. */
function eyebrow(slide, text, x, y, extra) {
	slide.addText(text, Object.assign({
		x, y, w: 1.972, h: 0.337, fontFace: BODY, fontSize: 14, color: LIME, valign: 'top',
	}, extra || {}));
}

/** Big Kanit headline; `lines` is a string or an array of lines. */
function heading(slide, lines, o) {
	const arr = [].concat(lines);
	slide.addText(arr.map((t, i) => ({ text: t, options: { breakLine: i < arr.length - 1 } })), Object.assign({
		w: 5.512, h: arr.length > 1 ? 1.313 : 0.707,
		fontFace: HEAD, fontSize: 36, color: WHITE, valign: 'top',
	}, o));
}

/** Body copy: 11pt Roboto, 150% leading. */
function body(slide, text, o) {
	slide.addText(text, Object.assign({
		fontFace: BODY, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top',
	}, o));
}

/** Bold sub-heading used above body copy. */
function subhead(slide, text, o) {
	slide.addText(text, Object.assign({
		h: 0.337, fontFace: BODY, fontSize: 14, bold: true, color: WHITE, valign: 'top',
	}, o));
}

/** Lime slanted tile carrying a small dark dumbbell icon. */
function iconTile(slide, x, y) {
	poly(slide, { x, y, w: 1.025, h: 0.685, pts: geo.parallelogram(1.025, 0.685, 0.33241), fill: { color: LIME } });
	dumbbells(slide, x + 0.2535, y + 0.2275, 0.518, DARK);
}

/** Slanted pill button. */
function pill(slide, text, x, y, o) {
	const s = Object.assign({ w: 1.575, h: 0.315, adj: 0.37185, fill: LIME, color: INK, size: 12 }, o || {});
	poly(slide, {
		x, y, w: s.w, h: s.h, flipH: true, shadow: CARD_SHADOW,
		pts: geo.parallelogram(s.w, s.h, s.adj),
		fill: s.fill ? { color: s.fill } : { color: DARK, transparency: 100 },
		line: s.line ? { color: s.line, width: 1 } : NO_LINE,
		text, textOpts: { fontSize: s.size, color: s.color },
	});
}

/**
 * Screen area of a device mock-up, standing in for the artwork of the original.
 * `o.clear` makes the glass translucent so background shapes show through, as
 * they do on the glossy device photos of the template.
 */
function screen(slide, x, y, w, h, o) {
	const s = o || {};
	slide.addText('[image]', {
		x, y, w, h, shape: s.radius ? 'roundRect' : 'rect', rectRadius: s.radius,
		fill: { color: '262626', transparency: s.clear || 0 }, line: NO_LINE,
		fontFace: BODY, fontSize: 11, color: '5C5C5C', align: 'center', valign: 'middle',
	});
}

/* Check mark used in the feature lists. */
const CHECK = [[539, 0], [234, 305], [119, 190], [0, 309], [115, 424], [234, 544], [354, 424], [658, 120]];
function checkMark(slide, x, y, w, h, color) {
	poly(slide, { x, y, w, h, pts: geo.frac(w, h, CHECK.map(p => [p[0] / 658, p[1] / 544])), fill: { color } });
}

/* Circled tick used on the pricing table. */
function tickCircle(slide, x, y, d, color) {
	slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: DARK, transparency: 100 }, line: { color, width: 0.75 } });
	poly(slide, {
		x: x + d * 0.24, y: y + d * 0.31, w: d * 0.52, h: d * 0.36,
		pts: geo.frac(d * 0.52, d * 0.36, [[0, 0.5], [0.16, 0.34], [0.4, 0.62], [0.86, 0], [1, 0.15], [0.4, 1]]),
		fill: { color },
	});
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

/** 1 & 19 & 30 — full-bleed cover slides. */
function coverSlide(pptx, title, size) {
	const s = pptx.addSlide();
	s.addShape('rect', { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: DARK, transparency: 30 }, line: NO_LINE });
	s.addText(title, {
		x: 1.425, y: size === 96 ? 2.892 : 2.959, w: 10.484, h: size === 96 ? 1.717 : 1.582,
		fontFace: HEAD, fontSize: size, color: WHITE, charSpacing: 3, align: 'center', valign: 'top',
	});
	s.addText('Gym & Fitness Presentation Template', {
		x: 3.945, y: 4.608, w: 5.443, h: 0.37,
		fontFace: BODY, fontSize: 16, color: WHITE, align: 'center', valign: 'top',
	});
	chrome(s);
	return s;
}

/** 2 — Welcome to Gym Workout */
function slide02(pptx) {
	const s = pptx.addSlide();
	poly(s, { x: 0.906, y: 3.764, w: 4.443, h: 3.736, pts: geo.parallelogram(4.443, 3.736, 0.73361), fill: { color: PANEL } });
	poly(s, {
		x: 0, y: 2.208, w: 4.444, h: 5.292, fill: { color: LIME },
		pts: geo.frac(4.444, 5.292, [[0.232, 0], [1, 0], [0.14, 1], [0, 1], [0, 0.27]]),
	});
	eyebrow(s, 'ABOUT US', 7.079, 2.043);
	heading(s, ['WELCOME TO', 'GYM WORKOUT'], { x: 7.079, y: 2.437, w: 5.1 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for. A wonderful serenity has taken possession of my entire soul, like these sweet',
		{ x: 7.079, y: 3.986, w: 4.726, h: 1.459 });
	pill(s, 'More Info', 7.216, 5.901, { color: '000000' });
	chrome(s);
}

/** 3 — About our gym workout */
function slide03(pptx) {
	const s = pptx.addSlide();
	poly(s, { x: 0, y: 1.441, w: 4.809, h: 6.059, pts: geo.parallelogram(4.809, 6.059, 0.52848), fill: { color: LIME } });
	eyebrow(s, 'ABOUT US', 7.076, 2.038);
	heading(s, ['ABOUT OUR', 'GYM WORKOUT'], { x: 7.076, y: 2.432, w: 5.1 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. am alone, and feel the charm of existence in this spot, which was created. A wonderful serenity',
		{ x: 7.076, y: 3.986, w: 4.571, h: 1.181 });
	subhead(s, 'Subtitle Gym & Fitness', { x: 7.076, y: 5.472, w: 3.539, fontSize: 12 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy',
		{ x: 7.076, y: 5.783, w: 4.321, h: 0.625 });
	poly(s, {
		x: 3.521, y: 4.896, w: 1.81, h: 1.163, shadow: CARD_SHADOW,
		pts: geo.parallelogram(1.81, 1.163, 0.44416), fill: { color: LIME },
	});
	s.addText('500+', { x: 3.723, y: 5.498, w: 1.263, h: 0.438, fontFace: BODY, fontSize: 20, color: INK, align: 'center', valign: 'top' });
	dumbbells(s, 4.156, 5.129, 0.668, DARK);
	chrome(s);
}

/** 4 — About program (right-hand lime wedge) */
function slide04(pptx) {
	const s = pptx.addSlide();
	poly(s, {
		x: 8.635, y: 0.915, w: 4.698, h: 6.585, flipV: true, fill: { color: LIME },
		pts: geo.frac(4.698, 6.585, [[0.678, 0], [1, 0], [1, 1], [0, 1]]),
	});
	poly(s, {
		x: 7.431, y: 5.122, w: 2.902, h: 1.464, shadow: CARD_SHADOW,
		pts: geo.parallelogram(2.902, 1.464, 0.5413), fill: { color: LIME }, line: { color: WHITE, width: 0.75 },
	});
	s.addText('800+', { x: 8.032, y: 5.426, w: 1.701, h: 0.572, fontFace: BODY, fontSize: 28, bold: true, color: INK, align: 'center', valign: 'top' });
	s.addText('Gym Workout', { x: 8.076, y: 6.023, w: 1.612, h: 0.286, fontFace: BODY, fontSize: 11, color: INK, align: 'center', valign: 'top' });
	eyebrow(s, 'ABOUT US', 1.155, 1.802);
	heading(s, ['ABOUT PROGRAM', 'GYM WORKOUT'], { x: 1.155, y: 2.196 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. am alone, and feel the charm of existence in this spot',
		{ x: 1.16, y: 3.759, w: 5.062, h: 0.903 });
	[['Body Fitness', 1.16], ['Fitness Health', 3.896]].forEach(([label, x]) => {
		checkMark(s, x, 5.062, 0.314, 0.216, LIME);
		subhead(s, label, { x: x + 0.389, y: 5.002, w: 1.937, fontSize: 12 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet', { x, y: 5.368, w: 2.326, h: 0.903 });
	});
	chrome(s, { stripes: false });
}

/** 5 — About program (two lime slabs, check list) */
function slide05(pptx) {
	const s = pptx.addSlide();
	poly(s, { x: 7.033, y: 0.677, w: 2.011, h: 3.073, pts: geo.parallelogram(2.011, 3.073, 0.35065), fill: { color: LIME } });
	poly(s, { x: 8.221, y: 3.75, w: 3.571, h: 3.75, pts: geo.parallelogram(3.571, 3.75, 0.23414), fill: { color: LIME } });
	poly(s, { x: 7.033, y: 0.677, w: 2.011, h: 3.073, pts: geo.parallelogram(2.011, 3.073, 0.35065), fill: { color: LIME } });
	poly(s, { x: 8.221, y: 3.75, w: 3.571, h: 3.75, pts: geo.parallelogram(3.571, 3.75, 0.23414), fill: { color: LIME } });
	eyebrow(s, 'ABOUT US', 1.162, 1.614);
	heading(s, ['ABOUT PROGRAM', 'GYM WORKOUT'], { x: 1.162, y: 2.009 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. am alone, and feel the charm of existence in this spot, which was created. A wonderful serenity has taken possession',
		{ x: 1.169, y: 3.516, w: 4.72, h: 1.181 });
	[['HIGH ACCESSIBILITY', 5.059, 1.348, 1.893], ['RENEWAL CONTRACT', 5.673, 1.348, 1.893], ['SPORT FITNESS', 6.287, 1.357, 1.902]].forEach(([label, y, cx, tx]) => {
		checkMark(s, cx, y + 0.001, 0.378, 0.314, LIME);
		s.addText(label, { x: tx, y, w: 3.046, h: 0.303, fontFace: BODY, fontSize: 12, color: WHITE, charSpacing: 3, valign: 'top' });
	});
	chrome(s, { stripes: false });
}

/** 6 — Design program (two icon rows) */
function slide06(pptx) {
	const s = pptx.addSlide();
	[3.4, 4.968].forEach((y, i) => {
		iconTile(s, 7.08, y + 0.139);
		subhead(s, 'Subtitle Gym Fitness', { x: 8.275, y, w: 3.603 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. ',
			{ x: 8.275, y: y + 0.365, w: i === 0 ? 3.89 : 3.904, h: 0.903 });
	});
	eyebrow(s, 'ABOUT US', 7.087, 1.489);
	heading(s, ['DESIGN PROGRAM', 'GYM WORKOUT'], { x: 7.087, y: 1.883 });
	chrome(s);
}

/** 7 — Design program (pentagon arrows + labelled rows) */
function slide07(pptx) {
	const s = pptx.addSlide();
	poly(s, { x: 0.018, y: 0, w: 6.301, h: 4.649, pts: geo.homePlate(6.301, 4.649, 0.31185), fill: { color: PANEL }, shadow: CARD_SHADOW });
	poly(s, { x: 0, y: 1.655, w: 6.467, h: 4.93, pts: geo.homePlate(6.467, 4.93, 0.30938), fill: { color: LIME }, shadow: CARD_SHADOW });
	eyebrow(s, 'ABOUT US', 7.48, 1.314);
	heading(s, 'DESIGN PROGRAM', { x: 7.48, y: 1.708 });
	[['BODY BUILDING', 2.73, 7.521, 3.125, 7.48], ['GUM TRAINING', 4.051, 7.521, 4.445, 7.48], ['CARDIO FITNESS', 5.372, 7.562, 5.766, 7.521]].forEach(([label, by, bx, ty, tx]) => {
		pill(s, label, bx, by, { w: 1.874, h: 0.324, size: 11 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my', { x: tx, y: ty, w: 4.619, h: 0.625 });
	});
	chrome(s);
}

/** 8 — Our workout strategy (three numbered cards) */
function slide08(pptx) {
	const s = pptx.addSlide();
	poly(s, {
		x: 6.859, y: 0, w: 6.201, h: 7.5, fill: { color: LIME },
		pts: geo.frac(6.201, 7.5, [[0.413, 0], [1, 0], [1, 1], [0, 1]]),
	});
	[['01', 6.418, 1.073], ['02', 5.739, 2.992], ['03', 5.124, 4.911]].forEach(([num, cx, cy]) => {
		poly(s, { x: cx, y: cy, w: 4.183, h: 1.685, pts: geo.parallelogram(4.183, 1.685, 0.3539), fill: { color: WHITE }, shadow: CARD_SHADOW });
		poly(s, { x: cx - 0.165, y: cy + 0.568, w: 0.887, h: 0.509, pts: geo.parallelogram(0.887, 0.509, 0.39341), fill: { color: LIME }, shadow: CARD_SHADOW });
		s.addText(num, { x: cx - 0.115, y: cy + 0.63, w: 0.78, h: 0.404, fontFace: BODY, fontSize: 18, bold: true, color: INK, align: 'center', valign: 'top' });
		subhead(s, 'Subtitle Strategy', { x: cx + 1.005, y: cy + 0.232, w: 2.344, color: INK });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings', { x: cx + 1.014, y: cy + 0.596, w: 2.464, h: 0.903, color: INK });
	});
	eyebrow(s, 'ABOUT US', 1.171, 2.046);
	heading(s, ['OUR WORKOUT', 'STRATEGY'], { x: 1.171, y: 2.44, w: 4.52 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for. A wonderful serenity has taken possession',
		{ x: 1.17, y: 3.997, w: 3.406, h: 1.736 });
	chrome(s, { stripes: false });
}

/** 9 — Wo we are */
function slide09(pptx) {
	const s = pptx.addSlide();
	const wedge = (w, h) => geo.frac(w, h, [[0, 0], [1, 0], [1, 1], [0.372, 1]]);
	poly(s, { x: 7.565, y: 3.75, w: 2.817, h: 3.75, pts: wedge(2.817, 3.75), fill: { color: PANEL } });
	poly(s, { x: 7.344, y: 1.397, w: 3.062, h: 3.978, pts: wedge(3.062, 3.978), fill: { color: LIME } });
	poly(s, { x: 7.565, y: 3.75, w: 2.817, h: 3.75, pts: wedge(2.817, 3.75), fill: { color: PANEL } });
	poly(s, { x: 7.344, y: 1.397, w: 3.062, h: 3.978, pts: wedge(3.062, 3.978), fill: { color: LIME } });
	eyebrow(s, 'ABOUT US', 1.168, 1.9);
	heading(s, ['WO WE ARE', 'GYM WORKOUT'], { x: 1.168, y: 2.294 });
	[3.956, 5.296].forEach((y, i) => {
		subhead(s, 'Subtitle Gym Fitness', { x: i ? 1.179 : 1.169, y, w: 3.603 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. ',
			{ x: i ? 1.179 : 1.169, y: y + 0.365, w: 5.115, h: 0.625 });
	});
	chrome(s);
}

/** 10 — Workout plan */
function slide10(pptx) {
	const s = pptx.addSlide();
	const wedge = (w, h) => geo.frac(w, h, [[0, 0], [1, 0], [1, 0.671], [0.444, 1]]);
	poly(s, { x: 5.642, y: 1.394, w: 4.823, h: 2.173, pts: wedge(4.823, 2.173), fill: { color: PANEL } });
	poly(s, { x: 4.648, y: 0, w: 6.449, h: 2.905, pts: wedge(6.449, 2.905), fill: { color: LIME } });
	poly(s, { x: 5.642, y: 1.394, w: 4.823, h: 2.173, pts: wedge(4.823, 2.173), fill: { color: PANEL } });
	poly(s, { x: 4.648, y: 0, w: 6.449, h: 2.905, pts: wedge(6.449, 2.905), fill: { color: LIME } });
	[['01', 1.155, 1.888, 1.985, 1.818], ['02', 1.18, 3.592, 2.01, 3.522], ['03', 1.155, 5.387, 1.985, 5.317]].forEach(([num, bx, by, tx, ty]) => {
		poly(s, { x: bx, y: by, w: 0.633, h: 0.4, pts: geo.parallelogram(0.633, 0.4, 0.35268), fill: { color: LIME }, shadow: CARD_SHADOW });
		s.addText(num, { x: bx, y: by + 0.04, w: 0.633, h: 0.337, fontFace: BODY, fontSize: 14, bold: true, color: INK, align: 'center', valign: 'top' });
		subhead(s, 'Subtitle Plan Here ', { x: tx, y: ty, w: 2.344 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which ', { x: tx, y: ty + 0.365, w: 3.117, h: 0.903 });
	});
	eyebrow(s, 'ABOUT US', 6.667, 4.396);
	heading(s, 'WORKOUT PLAN', { x: 6.667, y: 4.79 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. am alone, and feel the charm of existence in this spot',
		{ x: 6.667, y: 5.687, w: 5.181, h: 0.903 });
	chrome(s, { stripes: false });
}

/** 11 — Workout vision */
function slide11(pptx) {
	const s = pptx.addSlide();
	poly(s, {
		x: 0, y: 2.175, w: 4.476, h: 5.325, fill: { color: LIME },
		pts: geo.frac(4.476, 5.325, [[0, 0], [1, 0], [0.785, 1], [0, 1]]),
	});
	eyebrow(s, 'ABOUT US', 7.453, 2.377);
	heading(s, 'WORKOUT VISION', { x: 7.453, y: 2.771 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for. A wonderful serenity has taken possession',
		{ x: 7.453, y: 3.764, w: 4.726, h: 1.181 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, ',
		{ x: 7.453, y: 5.058, w: 4.726, h: 0.903 });
	chrome(s);
}

/** 12 — Workout mission */
function slide12(pptx) {
	const s = pptx.addSlide();
	poly(s, {
		x: 7.596, y: 3.75, w: 2.708, h: 3.75, fill: { color: PANEL },
		pts: geo.frac(2.708, 3.75, [[0, 0], [0.407, 0], [1, 1], [0.593, 1]]),
	});
	poly(s, {
		x: 6.369, y: 0, w: 5.618, h: 2.651, fill: { color: LIME },
		pts: geo.frac(5.618, 2.651, [[0, 0], [1, 0], [1, 1], [0.206, 1]]),
	});
	poly(s, {
		x: 7.596, y: 3.75, w: 2.708, h: 3.75, fill: { color: PANEL },
		pts: geo.frac(2.708, 3.75, [[0, 0], [0.407, 0], [1, 1], [0.593, 1]]),
	});
	poly(s, {
		x: 6.369, y: 0, w: 5.618, h: 2.651, fill: { color: LIME },
		pts: geo.frac(5.618, 2.651, [[0, 0], [1, 0], [1, 1], [0.206, 1]]),
	});
	eyebrow(s, 'ABOUT US', 1.166, 1.599);
	heading(s, 'WORKOUT MISSION', { x: 1.166, y: 1.993 });
	[[1.155, 3.212, 2.332, 3.087], [1.167, 4.441, 2.345, 4.316], [1.166, 5.722, 2.344, 5.597]].forEach(([ix, iy, tx, ty]) => {
		iconTile(s, ix, iy);
		subhead(s, 'Subtitle Mission ', { x: tx, y: ty, w: 2.344 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which ', { x: tx + 0.009, y: ty + 0.364, w: 4.089, h: 0.625 });
	});
	chrome(s, { stripes: false });
}

/** 13 — Our expert trainers (4 up) */
function slide13(pptx) {
	const s = pptx.addSlide();
	eyebrow(s, 'TEAM SLIDE', 5.681, 0.865, { align: 'center' });
	s.addText('OUR EXPERT TRAINERS', { x: 3.219, y: 1.259, w: 6.896, h: 0.64, fontFace: HEAD, fontSize: 32, color: WHITE, align: 'center', valign: 'top' });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. am alone, and feel the charm of existence in this spot',
		{ x: 2.812, y: 2.029, w: 7.72, h: 0.625, align: 'center' });
	[['Profession Trainers ', 0.737, 0.79], ['Professional Trainers', 3.562, 3.615], ['Professional Trainers', 6.386, 6.439], ['Professional Trainers', 9.211, 9.264]].forEach(([role, nx, rx], i) => {
		subhead(s, 'Member Name', { x: nx, y: i === 0 ? 5.946 : 5.945, w: 2.33, align: 'center' });
		s.addText(role, {
			x: rx, y: i === 0 ? 6.266 : 6.264, w: 2.224, h: 0.303,
			fontFace: BODY, fontSize: 9, italic: true, color: LIME, align: 'center', lineSpacingMultiple: 1.5, valign: 'top',
		});
	});
	chrome(s);
}

/** 14 — Our expert trainers (name plates) */
function slide14(pptx) {
	const s = pptx.addSlide();
	[[7.884, 4.407], [8.883, 0.367]].forEach(([x, y]) => {
		poly(s, { x, y, w: 2.747, h: 2.747, pts: geo.snip2Diag(2.747, 2.747, 0, 0.5), fill: { color: PANEL } });
	});
	[[7.484, 5.45, 7.807, 'JESSICA JANE'], [9.8, 0.93, 10.123, 'FERNANDES']].forEach(([px, py, tx, name]) => {
		poly(s, { x: px, y: py, w: 2.379, h: 1.13, pts: geo.snip2Diag(2.379, 1.13, 0, 0.31457), fill: { color: LIME } });
		s.addText(name, { x: tx, y: py + 0.233, w: 2.056, h: 0.337, fontFace: BODY, fontSize: 14, bold: true, color: INK, valign: 'top' });
		s.addText('CROSSFIT EXPERT', { x: tx, y: py + 0.566, w: 2.056, h: 0.348, fontFace: BODY, fontSize: 11, bold: true, italic: true, color: INK, lineSpacingMultiple: 1.5, valign: 'top' });
	});
	eyebrow(s, 'TEAM SLIDE', 1.178, 2.047);
	heading(s, ['OUR EXPERT', 'TRAINERS'], { x: 1.178, y: 2.441, w: 4.584 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for. A wonderful serenity has taken possession of my entire soul, like these sweet',
		{ x: 1.178, y: 3.99, w: 4.489, h: 1.459 });
	pill(s, 'More Info', 1.314, 5.904, { color: '000000' });
	chrome(s, { stripes: false });
}

/** 15 — Professional trainers (skill bars) */
function slide15(pptx) {
	const s = pptx.addSlide();
	gradPoly(s, { x: 4.867, y: 0, w: 6.006, h: 7.5, flipH: true, pts: geo.chevron(6.006, 7.5, 0.45234), bands: 96 });
	eyebrow(s, 'TEAM SLIDE', 1.167, 2.02);
	heading(s, ['PROFFESIONAL', 'TRAINERS'], { x: 1.167, y: 2.414, w: 4.043, h: 1.178, fontSize: 32 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. am alone, and feel the charm',
		{ x: 1.161, y: 3.887, w: 3.381, h: 1.181 });
	s.addText('Nathan T joe', { x: 9.713, y: 2.577, w: 2.884, h: 0.505, fontFace: HEAD, fontSize: 24, bold: true, color: LIME, valign: 'top' });
	s.addText('Professional Trainers', { x: 9.713, y: 3.098, w: 2.193, h: 0.303, fontFace: BODY, fontSize: 12, color: WHITE, valign: 'top' });
	// Skill rows: white slab, coloured percentage slab, label.
	[
		{ y: 3.768, bx: 8.822, px: 8.82, pw: 1.127, ph: 0.796, py: 3.762, pc: LIME, pct: '90%', label: 'Body Balacce', lx: 9.909, lw: 1.998, ty: 3.95, tx: 9.012 },
		{ y: 4.787, bx: 8.614, px: 8.612, pw: 1.113, ph: 0.801, py: 4.776, pc: LIME_D, pct: '85%', label: 'Cardio Bodystep', lx: 9.65, lw: 2.269, ty: 4.973, tx: 8.798 },
		{ y: 5.794, bx: 8.379, px: 8.377, pw: 1.13, ph: 0.79, py: 5.794, pc: LIME_DD, pct: '65%', label: 'Boxing & Run', lx: 9.644, lw: 1.921, ty: 5.971, tx: 8.56 },
	].forEach(r => {
		poly(s, { x: r.bx, y: r.y, w: 3.565, h: 0.79, pts: geo.parallelogram(3.565, 0.79, 0.25), fill: { color: WHITE }, shadow: CARD_SHADOW });
		poly(s, { x: r.px, y: r.py, w: r.pw, h: r.ph, pts: geo.parallelogram(r.pw, r.ph, 0.25), fill: { color: r.pc }, shadow: CARD_SHADOW });
		s.addText(r.pct, { x: r.tx, y: r.ty, w: 0.8, h: 0.404, fontFace: BODY, fontSize: 18, color: WHITE, align: 'center', valign: 'top' });
		s.addText(r.label, { x: r.lx, y: r.y + 0.254, w: r.lw, h: 0.303, fontFace: BODY, fontSize: 12, color: INK, align: 'right', charSpacing: 3, valign: 'top' });
	});
	// Thin progress bars under the body copy.
	[['Frofessional Trainers ', 5.412, 1.969, '60%'], ['Frofessional Trainers', 6.014, 2.52, '80%']].forEach(([label, y, fill, pct]) => {
		s.addText(label, { x: 1.245, y, w: 2.36, h: 0.226, fontFace: BODY, fontSize: 9, color: WHITE, valign: 'middle', margin: [0, 0.0394, 0, 0] });
		slideBar(s, 1.263, y + 0.279, 3.15, 0.079, PANEL);
		slideBar(s, 1.245, y + 0.279, fill, 0.079, LIME);
		s.addText(pct, { x: 4.546, y: y + 0.211, w: 0.594, h: 0.226, fontFace: BODY, fontSize: 9, color: WHITE, valign: 'middle' });
	});
	chrome(s);
}

function slideBar(s, x, y, w, h, color) {
	s.addShape('roundRect', { x, y, w, h, rectRadius: h / 2, fill: { color }, line: NO_LINE });
}

/** 16 — Workout services (2x2) */
function slide16(pptx) {
	const s = pptx.addSlide();
	eyebrow(s, 'SERVICES  SLIDE', 5.5, 1.074, { w: 2.351, align: 'center' });
	s.addText('WORKOUT SERVICES', { x: 3.219, y: 1.468, w: 6.896, h: 0.707, fontFace: HEAD, fontSize: 36, color: WHITE, align: 'center', valign: 'top' });
	[[1.64, 3.059, 2.817], [1.64, 4.979, 2.817], [7.261, 3.059, 8.438], [7.261, 4.979, 8.438]].forEach(([ix, iy, tx]) => {
		iconTile(s, ix, iy);
		subhead(s, 'Subtitle Services', { x: tx, y: iy - 0.125, w: 2.344 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which ', { x: tx + 0.009, y: iy + 0.239, w: 3.246, h: 0.903 });
	});
	chrome(s);
}

/** 17 — Best services */
function slide17(pptx) {
	const s = pptx.addSlide();
	poly(s, { x: 6.823, y: 0.915, w: 3.907, h: 2.364, flipH: true, pts: geo.parallelogram(3.907, 2.364, 0.50576), fill: { color: LIME } });
	eyebrow(s, 'SERVICES  SLIDE', 1.52, 1.554, { w: 2.351 });
	heading(s, 'BEST SERVICES', { x: 1.52, y: 1.948, w: 5.142 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. am alone, and feel the charm of existence in this spot,',
		{ x: 1.523, y: 2.836, w: 5.142, h: 0.903 });
	[[1.624, 4.394, 2.802], [1.615, 5.72, 2.793], [7.383, 4.394, 8.56], [7.374, 5.72, 8.551]].forEach(([ix, iy, tx]) => {
		iconTile(s, ix, iy);
		subhead(s, 'Subtitle Services', { x: tx, y: iy - 0.125, w: 2.344 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet', { x: tx + 0.009, y: iy + 0.239, w: 3.126, h: 0.625 });
	});
	chrome(s, { stripes: false });
}

/** 18 — Amazing best services */
function slide18(pptx) {
	const s = pptx.addSlide();
	poly(s, { x: 4.565, y: 0, w: 0.601, h: 2.165, flipH: true, pts: geo.parallelogram(0.601, 2.165, 0.57761), fill: { color: LIME } });
	[[8.243, 1.747, 9.468, 9.477], [8.531, 3.543, 9.756, 9.765], [8.775, 5.284, 10.001, 10.009]].forEach(([ix, iy, sx, bx]) => {
		iconTile(s, ix, iy);
		subhead(s, 'Subtitle Services', { x: sx, y: iy - 0.125, w: 2.344 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet', { x: bx, y: iy + 0.239, w: 2.335, h: 0.903 });
	});
	eyebrow(s, 'TEAM SLIDE', 1.168, 2.181);
	heading(s, ['AMAZING', 'BEST SERVICES'], { x: 1.168, y: 2.575, w: 3.998, h: 1.178, fontSize: 32 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for. A wonderful serenity has taken possession of my entire soul, like these sweet',
		{ x: 1.178, y: 3.958, w: 3.404, h: 2.014 });
	chrome(s);
}

/** 20 — Workout equipment (gallery) */
function slide20(pptx) {
	const s = pptx.addSlide();
	eyebrow(s, 'GALLERY  SLIDE', 5.5, 0.932, { w: 2.351, align: 'center' });
	s.addText('WORKOUT EQUIPMENT', { x: 3.219, y: 1.326, w: 6.896, h: 0.707, fontFace: HEAD, fontSize: 36, color: WHITE, align: 'center', valign: 'top' });
	[['Fitness Schedule', 1.537, 1.086, 5.587, 5.951], ['Superior Facilities', 5.162, 4.711, 5.596, 5.961], ['Muscle Gym', 8.805, 8.353, 5.587, 5.951]].forEach(([label, tx, bx, ty, by]) => {
		subhead(s, label, { x: tx, y: ty, w: 2.344, align: 'center' });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet', { x: bx, y: by, w: 3.246, h: 0.625, align: 'center' });
	});
	// Decorative chevrons.
	[[1.242, 3.48, 180, true, false], [6.478, 2.346, 270, true, true], [11.467, 4.905, 180, false, false]].forEach(([x, y, rot, fh, fv]) => {
		poly(s, { x, y, w: 0.408, h: 0.54, rotate: rot, flipH: fh, flipV: fv, pts: geo.chevron(0.408, 0.54, 0.5), fill: { color: LIME } });
	});
	chrome(s);
}

/** 21 — Grup workout classes */
function slide21(pptx) {
	const s = pptx.addSlide();
	eyebrow(s, 'GALLERY  SLIDE', 1.169, 1.402, { w: 2.351 });
	heading(s, ['GRUP WORKOUT', 'CLASSES'], { x: 1.169, y: 1.796, w: 5.083 });
	[3.553, 5.12].forEach(iy => {
		iconTile(s, 1.155, iy);
		subhead(s, 'Subtitle Gym Fitness', { x: 2.349, y: iy - 0.14, w: 3.603 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. ', { x: 2.349, y: iy + 0.225, w: 3.603, h: 0.903 });
	});
	[[6.538, 1.152], [11.202, 5.061]].forEach(([x, y]) => {
		poly(s, { x, y, w: 0.67, h: 1.564, pts: geo.parallelogram(0.67, 1.564, 0.57761), fill: { color: LIME } });
	});
	chrome(s);
}

/** 22 — Our portfolio */
function slide22(pptx) {
	const s = pptx.addSlide();
	gradPoly(s, { x: 6.667, y: 0, w: 4.627, h: 6.336, flipH: true, pts: geo.chevron(4.627, 6.336, 0.42714), bands: 90 });
	gradPoly(s, { x: 1.155, y: 2.206, w: 5.512, h: 5.294, flipH: true, up: true, pts: geo.chevron(5.512, 5.294, 0.37466), bands: 90 });
	eyebrow(s, 'GALLERY  SLIDE', 7.906, 1.405, { w: 2.351 });
	heading(s, 'OUR PORTFOLIO', { x: 7.906, y: 1.799, w: 5.083 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. ', { x: 7.906, y: 2.619, w: 4.027, h: 0.903 });
	chrome(s);
}

/** 23 — App phone mockup */
function slide23(pptx) {
	const s = pptx.addSlide();
	gradPoly(s, { x: 1.177, y: 0, w: 5.49, h: 7.5, flipH: true, up: true, pts: geo.chevron(5.49, 7.5, 0.36037), bands: 96 });
	phoneMockup(s, 0.995, 1.226, 2.644, 5.102);
	phoneMockup(s, 4.023, 1.797, 2.644, 5.102);
	eyebrow(s, 'DEVICE SLIDE', 7.476, 2.045);
	heading(s, ['APP PHONE', 'MOCKUP SLIDE'], { x: 7.476, y: 2.44, w: 4.681 });
	[4.094, 5.414].forEach(y => {
		speakerIcon(s, 7.59, y + 0.089);
		subhead(s, 'Description mockup', { x: 8.24, y, w: 3.179 });
		body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings', { x: 8.244, y: y + 0.366, w: 3.55, h: 0.625 });
	});
	chrome(s);
}

/** Stylised phone stand-in (replaces the photo mock-up of the original). */
function phoneMockup(s, x, y, w, h) {
	screen(s, x + 0.075, y + 0.075, w - 0.15, h - 0.15, { radius: 0.25, clear: 90 });
	// Bezel drawn as a stroke so the glass shows the artwork behind it.
	s.addShape('roundRect', { x, y, w, h, rectRadius: 0.3, fill: { color: DARK, transparency: 100 }, line: { color: '9AA0A6', width: 3 } });
	s.addShape('roundRect', { x: x + w / 2 - 0.42, y: y + 0.075, w: 0.84, h: 0.16, rectRadius: 0.08, fill: { color: '2E3135' }, line: NO_LINE });
}

/** Rounded lime button carrying a small speaker glyph. */
function speakerIcon(s, x, y) {
	s.addShape('roundRect', { x, y, w: 0.435, h: 0.435, rectRadius: 0.127, fill: { color: LIME }, line: NO_LINE, shadow: CARD_SHADOW });
	poly(s, {
		x: x + 0.093, y: y + 0.125, w: 0.15, h: 0.185, fill: { color: DARK },
		pts: geo.frac(0.15, 0.185, [[0, 0.3], [0.4, 0.3], [1, 0], [1, 1], [0.4, 0.7], [0, 0.7]]),
	});
	[[0.26, 0.16, 0.05], [0.30, 0.125, 0.07]].forEach(([dx, dy, r]) => {
		s.addShape('ellipse', { x: x + dx, y: y + dy, w: r, h: r * 2, fill: { color: DARK, transparency: 100 }, line: { color: DARK, width: 0.75 } });
	});
}

/** 24 — Website laptop mockup */
function slide24(pptx) {
	const s = pptx.addSlide();
	poly(s, {
		x: 8.844, y: 0, w: 4.49, h: 7.5, flipV: true, fill: { color: LIME },
		pts: geo.frac(4.49, 7.5, [[0.678, 0], [1, 0], [1, 1], [0, 1]]),
	});
	laptopMockup(s, 6.625, 1.402, 6.708, 5.125);
	eyebrow(s, 'DEVICE SLIDE', 1.168, 1.638);
	heading(s, ['WEBSITE LAPTOP', 'MOCKUP SLIDE'], { x: 1.168, y: 2.032 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for',
		{ x: 1.161, y: 3.493, w: 4.381, h: 1.181 });
	[['20', 'K', 'Impressive', 1.158], ['45', 'K', 'Professional', 3.377]].forEach(([num, unit, label, x]) => {
		s.addText([
			{ text: num, options: { fontSize: 36, bold: true, color: LIME } },
			{ text: unit, options: { fontSize: 20, bold: true, color: LIME } },
		], { x, y: 4.922, w: 1.483, h: 0.707, fontFace: BODY, paraSpaceBefore: 12, valign: 'top' });
		s.addText(label, { x, y: 5.629, w: 1.985, h: 0.303, fontFace: BODY, fontSize: 12, color: WHITE, valign: 'top' });
		s.addText('Lorem ipsum dolor', { x, y: 5.945, w: 1.985, h: 0.31, fontFace: BODY, fontSize: 10.5, color: WHITE, lineSpacingMultiple: 1.3, paraSpaceBefore: 12, valign: 'top' });
	});
	chrome(s, { stripeColor: DARK });
}

/**
 * Stylised open laptop (replaces the photo mock-up of the original). The lid
 * runs off the right edge of the slide exactly as the source photo does.
 */
function laptopMockup(s, x, y, w, h) {
	const lidH = h * 0.92;
	s.addShape('roundRect', { x: x + 0.81, y, w: w - 0.81, h: lidH, rectRadius: 0.09, fill: { color: '0A0A0A' }, line: { color: 'A9ADB2', width: 2 } });
	screen(s, x + 1.05, y + 0.31, w - 1.05, lidH - 0.62);
	s.addShape('roundRect', { x, y: y + lidH + 0.1, w, h: 0.28, rectRadius: 0.1, fill: { color: 'E9EBED' }, line: NO_LINE });
	s.addShape('roundRect', { x: x + 2.2, y: y + lidH + 0.11, w: 0.85, h: 0.05, rectRadius: 0.025, fill: { color: 'AFB3B8' }, line: NO_LINE });
}

/** 25 — Website desktop mockup */
function slide25(pptx) {
	const s = pptx.addSlide();
	poly(s, { x: 7.461, y: 2.292, w: 5.873, h: 5.208, flipH: true, pts: geo.rtTriangle(5.873, 5.208), fill: { color: LIME } });
	desktopMockup(s, 6.681, 1.036, 5.868, 5.868);
	eyebrow(s, 'DEVICE SLIDE', 1.168, 1.763);
	heading(s, ['WEBSITE DESKTOP', 'MOCKUP SLIDE'], { x: 1.168, y: 2.157 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for. A wonderful serenity has taken possession of my entire soul, like these sweet',
		{ x: 1.176, y: 3.764, w: 4.726, h: 1.459 });
	pill(s, 'More Info', 1.298, 5.706);
	pill(s, 'Website', 3.069, 5.706, { fill: null, line: LIME, color: WHITE });
	chrome(s);
}

/** Stylised all-in-one desktop with keyboard and mouse (replaces a photo). */
function desktopMockup(s, x, y, w, h) {
	const mx = x + 0.57, mw = 4.77, my = y + 0.66, mh = 2.94; // display panel
	screen(s, mx + 0.11, my + 0.15, mw - 0.22, mh - 0.28, { clear: 92 });
	s.addShape('rect', { x: mx, y: my, w: mw, h: mh, fill: { color: DARK, transparency: 100 }, line: { color: '0A0A0A', width: 10 } });
	s.addShape('rect', { x: mx + 0.015, y: my + mh, w: mw - 0.03, h: 0.55, fill: { color: 'C6C9CD' }, line: NO_LINE });
	poly(s, {
		x: mx + 1.77, y: my + mh + 0.55, w: 1.6, h: 0.57, fill: { color: 'B7BABF' },
		pts: geo.frac(1.6, 0.57, [[0.28, 0], [0.72, 0], [1, 1], [0, 1]]),
	});
	s.addShape('roundRect', { x: x + 1.55, y: y + 4.51, w: 3.15, h: 0.3, rectRadius: 0.06, fill: { color: 'E4E6E9' }, line: NO_LINE });
	s.addShape('ellipse', { x: x + 4.66, y: y + 4.56, w: 0.89, h: 0.22, fill: { color: 'E4E6E9' }, line: NO_LINE });
}

/** 26 — Infographic slide (four slanted cards) */
function slide26(pptx) {
	const s = pptx.addSlide();
	eyebrow(s, 'INFOGRAPHIC', 5.5, 0.932, { w: 2.351, align: 'center' });
	s.addText('INFOGRAPHIC SLIDE', { x: 3.219, y: 1.326, w: 6.896, h: 0.707, fontFace: HEAD, fontSize: 36, color: WHITE, align: 'center', valign: 'top' });
	[
		{ x: 1.069, y: 2.629, card: LIME, dot: DARK, mark: WHITE, text: INK },
		{ x: 3.894, y: 3.181, card: PANEL, dot: LIME, mark: DARK, text: WHITE },
		{ x: 6.586, y: 2.629, card: LIME, dot: DARK, mark: WHITE, text: INK },
		{ x: 9.412, y: 3.181, card: PANEL, dot: LIME, mark: DARK, text: WHITE },
	].forEach(c => {
		poly(s, { x: c.x, y: c.y, w: 2.841, h: 3.081, flipH: true, pts: geo.parallelogram(2.841, 3.081, 0.13428), fill: { color: c.card }, shadow: CARD_SHADOW });
		s.addShape('ellipse', { x: c.x + 0.988, y: c.y + 0.357, w: 0.762, h: 0.762, fill: { color: c.dot }, line: NO_LINE, shadow: CARD_SHADOW });
		dumbbells(s, c.x + 1.11, c.y + 0.623, 0.518, c.mark);
		s.addText('Description Text ', { x: c.x + 0.456, y: c.y + 1.392, w: 2.049, h: 0.303, fontFace: BODY, fontSize: 12, bold: true, color: c.text, align: 'center', valign: 'top' });
		s.addText('A wonderful serenity has taken possession of my entire soul, like these sweet', {
			x: c.x + 0.395, y: c.y + 1.765, w: 2.171, h: 0.83,
			fontFace: BODY, fontSize: 10, color: c.text, align: 'center', lineSpacingMultiple: 1.5, valign: 'top',
		});
	});
	chrome(s);
}

/** 27 — Infographic step (rising arrows) */
function slide27(pptx) {
	const s = pptx.addSlide();
	[
		{ x: 0.905, y: 3.572, num: '01', head: LIME, from: LIME_L, to: LIME, stop: 0.33 },
		{ x: 3.202, y: 2.774, num: '02', head: PANEL, from: SILVER, to: PANEL, stop: 0 },
		{ x: 5.499, y: 1.944, num: '03', head: LIME, from: LIME_L, to: LIME, stop: 0 },
		{ x: 7.796, y: 1.087, num: '04', head: PANEL, from: SILVER, to: PANEL, stop: 0 },
	].forEach(a => {
		gradBar(s, { x: a.x, y: a.y + 1.119, w: 3.07, h: 0.615, rotate: 315, from: a.from, to: a.to, stop: a.stop, bands: 56 });
		s.addShape('roundRect', { x: a.x + 1.032, y: a.y, w: 1.932, h: 0.615, rectRadius: 0.3075, fill: { color: a.head }, line: NO_LINE });
		s.addShape('roundRect', { x: a.x + 1.691, y: a.y + 0.658, w: 1.932, h: 0.615, rotate: 90, rectRadius: 0.3075, fill: { color: a.head }, line: NO_LINE });
		s.addText(a.num, { x: a.x + 2.078, y: a.y, w: 0.825, h: 0.503, fontFace: BODY, fontSize: 20, bold: true, color: WHITE, align: 'center', lineSpacingMultiple: 1.3, valign: 'top' });
	});
	[[3.131, 5.976, 2.269], [5.547, 4.94, 2.269], [8.022, 4.152, 2.308], [10.204, 3.19, 2.236]].forEach(([x, y, bw]) => {
		s.addText('Step Infographic', { x, y, w: 2.553, h: 0.303, fontFace: BODY, fontSize: 12, bold: true, color: WHITE, valign: 'top' });
		s.addText('A wonderful serenity has taken possession of my entire', {
			x: x + 0.004, y: y + 0.311, w: bw, h: 0.58,
			fontFace: BODY, fontSize: 10, color: WHITE, lineSpacingMultiple: 1.5, valign: 'top',
		});
	});
	eyebrow(s, 'INFOGRAPHIC', 1.169, 1.261, { w: 2.351 });
	heading(s, ['INFOGRAPHIC ', 'STEP'], { x: 1.169, y: 1.655, w: 3.97 });
	chrome(s);
}

/** 28 — Pricing table */
function slide28(pptx) {
	const s = pptx.addSlide();
	eyebrow(s, 'TABLES PRICING', 5.5, 0.932, { w: 2.351, align: 'center' });
	s.addText('YOUR PRICING PLAN', { x: 3.219, y: 1.326, w: 6.896, h: 0.707, fontFace: HEAD, fontSize: 36, color: WHITE, align: 'center', valign: 'top' });
	[
		{ x: 1.641, y: 2.418, h: 4.32, card: PANEL, title: 'Medium Business', tw: 2.208, tx: 2.081, price: '15', px: 2.29, py: 3.355, priceColor: GRAY, rowY: 4.069, rowX: 2.075, textColor: WHITE, tick: GRAY, btnX: 2.48, btnY: 6.067, btnFill: LIME, btnColor: INK },
		{ x: 5.164, y: 2.418, h: 4.334, card: LIME, title: 'Premiun Business', tw: 2.362, tx: 5.511, price: '30', px: 5.799, py: 3.321, priceColor: INK, rowY: 4.034, rowX: 5.585, textColor: INK, tick: DARK, btnX: 5.989, btnY: 6.033, btnFill: DARK, btnColor: WHITE },
		{ x: 8.66, y: 2.432, h: 4.32, card: PANEL, title: 'Pro Business', tw: 1.941, tx: 9.234, price: '40', px: 9.309, py: 3.337, priceColor: GRAY, rowY: 4.05, rowX: 9.095, textColor: WHITE, tick: GRAY, btnX: 9.499, btnY: 6.048, btnFill: LIME, btnColor: INK },
	].forEach(p => {
		s.addShape('roundRect', { x: p.x, y: p.y, w: 3.088, h: p.h, rectRadius: 0.232, fill: { color: p.card }, line: NO_LINE });
		s.addText(p.title, { x: p.tx, y: p.y + 0.339, w: p.tw, h: 0.337, fontFace: BODY, fontSize: 14, bold: true, color: p.textColor, align: 'center', valign: 'top' });
		s.addText([
			{ text: '$', options: { fontSize: 36, bold: true, color: p.priceColor, baseline: 600 } },
			{ text: p.price, options: { fontSize: 36, bold: true, color: p.priceColor } },
		], { x: p.px, y: p.py, w: 1.79, h: 0.598, fontFace: BODY, align: 'center', lineSpacingMultiple: 0.8, valign: 'top' });
		for (let i = 0; i < 4; i++) {
			const ry = p.rowY + i * 0.4505;
			tickCircle(s, p.rowX, ry + 0.117, 0.173, p.tick);
			s.addText('A wonderful serenity has', { x: p.rowX + 0.233, y: ry, w: 2.208, h: 0.322, fontFace: BODY, fontSize: 11, color: p.textColor, lineSpacingMultiple: 1.3, valign: 'top' });
		}
		s.addText('Read More', {
			x: p.btnX, y: p.btnY, w: 1.41, h: 0.455, shape: 'roundRect', rectRadius: 0.058,
			fill: { color: p.btnFill }, line: NO_LINE, shadow: CARD_SHADOW,
			fontFace: BODY, fontSize: 11, color: p.btnColor, align: 'center', valign: 'middle',
		});
	});
	chrome(s);
}

/** 29 — Get it touch (contact) */
function slide29(pptx) {
	const s = pptx.addSlide();
	const banner = (w, h) => geo.frac(w, h, [[0, 0], [1, 0], [1, 0.576], [0.888, 1], [0.029, 0.667]]);
	poly(s, { x: 4.309, y: 0.915, w: 9.025, h: 3.335, pts: banner(9.025, 3.335), fill: { color: PANEL } });
	poly(s, { x: 2.779, y: 0.642, w: 9.627, h: 3.108, pts: banner(9.627, 3.108), fill: { color: LIME } });
	poly(s, { x: 4.309, y: 0.915, w: 9.025, h: 3.335, pts: banner(9.025, 3.335), fill: { color: PANEL } });
	poly(s, { x: 2.779, y: 0.642, w: 9.627, h: 3.108, pts: banner(9.627, 3.108), fill: { color: LIME } });
	eyebrow(s, 'CONTACT US', 1.169, 3.771);
	heading(s, 'GET IT TOUCH', { x: 1.169, y: 4.165, w: 5.345 });
	body(s, 'A wonderful serenity has taken possession of my entire soul, like these sweet mornings of spring which I enjoy with my whole heart. I am alone, and feel the charm of existence in this spot, which was created for. ',
		{ x: 1.169, y: 5.065, w: 4.434, h: 1.181 });
	globeIcon(s, 7.26, 4.631);
	phoneIcon(s, 7.26, 5.167);
	pinIcon(s, 7.26, 5.789);
	[['contact@website.com', 4.544, 7.745, 3.398], ['(000) 1234 5678 45', 5.121, 7.745, 3.398], ['500 Random street, States, 2290', 5.8, 7.73, 3.606]].forEach(([t, y, x, w]) => {
		s.addText(t, { x, y, w, h: 0.382, fontFace: BODY, fontSize: 14, color: WHITE, lineSpacingMultiple: 1.3, paraSpaceBefore: 6, valign: 'top' });
	});
	chrome(s, { stripes: false });
}

function globeIcon(s, x, y) {
	const d = 0.29;
	s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: DARK, transparency: 100 }, line: { color: WHITE, width: 1 } });
	s.addShape('ellipse', { x: x + d * 0.32, y, w: d * 0.36, h: d, fill: { color: DARK, transparency: 100 }, line: { color: WHITE, width: 0.75 } });
	s.addShape('line', { x, y: y + d / 2, w: d, h: 0, line: { color: WHITE, width: 0.75 } });
}

function phoneIcon(s, x, y) {
	const d = 0.29;
	poly(s, {
		x, y, w: d, h: d, fill: { color: WHITE },
		pts: geo.frac(d, d, [[0.06, 0], [0.36, 0], [0.46, 0.28], [0.32, 0.42], [0.58, 0.7], [0.72, 0.55], [1, 0.66], [1, 0.95],
			[0.82, 1], [0.42, 0.86], [0.13, 0.56], [0, 0.18]]),
	});
}

function pinIcon(s, x, y) {
	s.addShape('teardrop', { x, y, w: 0.276, h: 0.276, rotate: 135, fill: { color: DARK, transparency: 100 }, line: { color: WHITE, width: 1 } });
	s.addShape('ellipse', { x: x + 0.093, y: y + 0.083, w: 0.09, h: 0.09, fill: { color: DARK, transparency: 100 }, line: { color: WHITE, width: 1 } });
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */
function build() {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'WIDE', width: SLIDE_W, height: SLIDE_H });
	pptx.layout = 'WIDE';
	pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
	pptx.title = 'WORKOUT — Gym & Fitness Presentation Template';

	pptx.defineSlideMaster({ title: 'BASE', background: { color: DARK } });
	const addSlide = pptx.addSlide.bind(pptx);
	pptx.addSlide = () => addSlide({ masterName: 'BASE' });

	coverSlide(pptx, 'WORKOUT', 96);
	slide02(pptx);
	slide03(pptx);
	slide04(pptx);
	slide05(pptx);
	slide06(pptx);
	slide07(pptx);
	slide08(pptx);
	slide09(pptx);
	slide10(pptx);
	slide11(pptx);
	slide12(pptx);
	slide13(pptx);
	slide14(pptx);
	slide15(pptx);
	slide16(pptx);
	slide17(pptx);
	slide18(pptx);
	coverSlide(pptx, 'BREAK SLIDE', 88);
	slide20(pptx);
	slide21(pptx);
	slide22(pptx);
	slide23(pptx);
	slide24(pptx);
	slide25(pptx);
	slide26(pptx);
	slide27(pptx);
	slide28(pptx);
	slide29(pptx);
	coverSlide(pptx, 'THANK YOU', 88);

	return pptx.writeFile({ fileName: path.join(__dirname, '16c18c51-bb9e-4a05-9aed-df3581fc82af_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
