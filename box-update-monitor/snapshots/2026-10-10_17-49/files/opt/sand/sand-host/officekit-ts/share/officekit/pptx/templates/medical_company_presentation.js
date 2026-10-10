/**
 * "MEDICAL" presentation template — recreated with pptxgenjs.
 *
 * Raster photos of the original deck are replaced by flat grey "[image]"
 * placeholder rectangles of the same position and size.
 *
 *   node 0658b937-7af8-446c-9851-b46f924081e5_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ *
 * Theme
 * ------------------------------------------------------------------ */

const GREEN = '02D090'; // accent1
const AQUA = '65D3DA'; // accent2
const BLUE = '0164FF'; // accent3
const TEAL = '28989F'; // accent4
const BLACK = '000000'; // tx1
const WHITE = 'FFFFFF'; // bg1
const LT2 = 'E7E6E6'; // bg2
const PHOTO = 'B9BABC'; // placeholder for raster images
const PHOTO_TX = '6E7073';

const SANS = 'Poppins';
const SANS_M = 'Poppins Medium';
const SANS_SB = 'Poppins SemiBold';
const SANS_T = 'Poppins Thin';
const INTER = 'Inter';
const INTER_M = 'Inter Medium';
const INTER_SB = 'Inter SemiBold';
const INTER_T = 'Inter Thin';

const CARD_R = 0.1394; // roundRect adj 6453 on a 2.158" high card
const PILL_R = 0.185; // roundRect adj 50000 on a 0.37" high pill

/* ------------------------------------------------------------------ *
 * Small helpers
 * ------------------------------------------------------------------ */

const mix = (a, b, t) => {
	const p = (i) => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t);
	return [p(0), p(2), p(4)].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
};

/**
 * Plain text box.  PowerPoint text boxes are top-anchored and use insets of
 * 0.1" horizontally / 0.05" vertically; `margin` is [left, right, bottom, top].
 */
function tx(slide, text, o) {
	slide.addText(text, Object.assign({ valign: 'top', margin: [7.2, 7.2, 3.6, 3.6], isTextBox: true }, o));
}

function rect(slide, x, y, w, h, o) {
	slide.addShape('rect', Object.assign({ x, y, w, h }, o));
}

function roundRect(slide, x, y, w, h, o) {
	slide.addShape('roundRect', Object.assign({ x, y, w, h, rectRadius: CARD_R }, o));
}

function ellipse(slide, x, y, w, h, o) {
	slide.addShape('ellipse', Object.assign({ x, y, w, h }, o));
}

function hLine(slide, x, y, w, color, width, transparency) {
	slide.addShape('line', { x, y, w, h: 0, line: { color, width, transparency } });
}

/* ------------------------------------------------------------------ *
 * Recurring decorations
 * ------------------------------------------------------------------ */

/**
 * The big two-tone ring ("Graphic 22"): a 34.8%-thick donut whose fill
 * fades from green at the bottom to teal at the top.  pptxgenjs has no
 * gradient fill, so the ring is built from angular slices.
 */
function ring(slide, x, y, size) {
	const SLICES = 20;
	for (let i = 0; i < SLICES; i++) {
		const a0 = (360 / SLICES) * i;
		const a1 = a0 + 360 / SLICES;
		const mid = ((a0 + a1) / 2) * (Math.PI / 180); // 0 = 3 o'clock, growing clockwise
		slide.addShape('blockArc', {
			x, y, w: size, h: size,
			fill: { color: mix(GREEN, TEAL, (1 - Math.sin(mid)) / 2) },
			angleRange: [a0, a1 + 0.6], // overlap slightly so no hairline shows between slices
			arcThicknessRatio: 0.3479,
		});
	}
}

/** Cluster of four "+" marks used as confetti all over the deck. */
const PLUSES = [
	[0.230, 0.000, 0.287, 0.257, BLUE, 0],
	[0.759, 0.271, 0.198, 0.177, GREEN, 0],
	[0.074, 0.557, 0.198, 0.177, GREEN, 0],
	[0.586, 0.735, 0.287, 0.253, BLUE, 65],
];

function plusCluster(slide, x, y) {
	PLUSES.forEach(([dx, dy, w, h, color, transparency]) => {
		slide.addShape('mathPlus', { x: x + dx, y: y + dy, w, h, fill: { color, transparency } });
	});
}

/** Blue pill button + label used as a footer badge. */
function pill(slide, x, y, label) {
	slide.addShape('roundRect', { x, y, w: 1.871, h: 0.37, fill: { color: BLUE }, rectRadius: PILL_R });
	tx(slide, label || 'General Medical', {
		x: x + 0.205, y: y + 0.062, w: 1.375, h: 0.269,
		fontFace: SANS_SB, fontSize: 10, color: WHITE, transparency: 12, align: 'center',
	});
}

function siteUrl(slide, x, y, color) {
	tx(slide, 'www.medical.com', {
		x, y, w: 3.773, h: 0.404, fontFace: SANS_T, fontSize: 18, charSpacing: 6, color: color || BLACK,
	});
}

/** Flat stand-in for a photograph. */
function photo(slide, x, y, w, h, o) {
	const opt = o || {};
	slide.addShape(opt.shape || 'rect', Object.assign(
		{ x, y, w, h, fill: { color: PHOTO } },
		opt.shape === 'roundRect' ? { rectRadius: CARD_R } : {}
	));
	if (opt.label !== false) {
		tx(slide, '[image]', {
			x, y: y + h / 2 - 0.16, w, h: 0.32,
			fontFace: INTER, fontSize: 11, color: PHOTO_TX, align: 'center',
		});
	}
}

/**
 * custGeom helper — `pts` is a compact list where a point is either
 * [x, y] (line), [x, y, x1, y1, x2, y2] (cubic) or the string 'close'.
 */
function poly(slide, x, y, w, h, fill, pts, extra) {
	const points = pts.map((p, i) => {
		if (p === 'close') return { close: true };
		if (p.length === 6) return { x: p[0], y: p[1], curve: { type: 'cubic', x1: p[2], y1: p[3], x2: p[4], y2: p[5] } };
		return i === 0 ? { x: p[0], y: p[1], moveTo: true } : { x: p[0], y: p[1] };
	});
	slide.addShape('custGeom', Object.assign({ x, y, w, h, fill, points }, extra));
}

/* ------------------------------------------------------------------ *
 * Icon glyphs (line-art icons of the original, drawn with native shapes)
 * ------------------------------------------------------------------ */

function icon(slide, kind, x, y, w, h, color) {
	const lw = Math.max(0.75, w * 3.2); // stroke weight, in points
	const L = { color, width: lw };
	const NO = { type: 'none' };
	const px = (u) => x + w * u;
	const py = (v) => y + h * v;

	if (kind === 'heart') {
		slide.addShape('heart', { x, y: y - h * 0.06, w, h: h * 1.06, fill: NO, line: L });
		// ECG zig-zag across the middle
		poly(slide, x, y, w, h, NO, [
			[w * 0.06, h * 0.46], [w * 0.3, h * 0.46], [w * 0.42, h * 0.72],
			[w * 0.56, h * 0.28], [w * 0.66, h * 0.46], [w * 0.94, h * 0.46],
		], { line: L });
	} else if (kind === 'tooth') {
		poly(slide, x, y, w, h, NO, [
			[w * 0.5, h * 0.06, w * 0.2, -h * 0.06, w * 0.32, h * 0.02],
			[w * 0.96, h * 0.34, w * 0.86, h * 0.02, w * 0.96, h * 0.14],
			[w * 0.72, h * 0.96, w * 0.96, h * 0.62, w * 0.84, h * 0.96],
			[w * 0.58, h * 0.58, w * 0.62, h * 0.96, w * 0.62, h * 0.62],
			[w * 0.42, h * 0.58, w * 0.54, h * 0.5, w * 0.46, h * 0.5],
			[w * 0.28, h * 0.96, w * 0.38, h * 0.62, w * 0.38, h * 0.96],
			[w * 0.04, h * 0.34, w * 0.16, h * 0.96, w * 0.04, h * 0.62],
			[w * 0.5, h * 0.06, w * 0.04, h * 0.14, w * 0.2, h * -0.02],
			'close',
		], { line: L });
	} else if (kind === 'stetho') {
		// chest piece + tubing + earpiece
		slide.addShape('roundRect', {
			x: px(0.02), y: py(0.10), w: w * 0.44, h: h * 0.36, rectRadius: w * 0.09, fill: NO, line: L,
		});
		slide.addShape('line', { x: px(0.24), y: py(0.10), w: 0, h: h * 0.36, line: L });
		poly(slide, x, y, w, h, NO, [
			[w * 0.24, h * 0.46], [w * 0.24, h * 0.72, w * 0.24, h * 0.62, w * 0.34, h * 0.80],
			[w * 0.70, h * 0.50, w * 0.60, h * 0.84, w * 0.70, h * 0.72],
			[w * 0.70, h * 0.34],
		], { line: L });
		ellipse(slide, px(0.60), py(0.10), w * 0.24, h * 0.24, { fill: NO, line: L });
	} else if (kind === 'caduceus') {
		ellipse(slide, x, y, w, h, { fill: NO, line: L });
		slide.addShape('line', { x: px(0.5), y: py(0.06), w: 0, h: h * 0.88, line: L });
		[0.26, 0.70].forEach((v) => {
			ellipse(slide, px(0.34), py(v), w * 0.14, h * 0.14, { fill: NO, line: L });
			ellipse(slide, px(0.52), py(v), w * 0.14, h * 0.14, { fill: NO, line: L });
		});
	} else if (kind === 'medkit') {
		slide.addShape('roundRect', {
			x, y: py(0.16), w, h: h * 0.68, rectRadius: w * 0.13, fill: NO, line: L,
		});
		slide.addShape('mathPlus', { x: px(0.3), y: py(0.36), w: w * 0.4, h: h * 0.28, fill: { color } });
	} else if (kind === 'ambulance') {
		slide.addShape('roundRect', {
			x, y: py(0.20), w: w * 0.96, h: h * 0.48, rectRadius: w * 0.1, fill: NO, line: L,
		});
		slide.addShape('mathPlus', { x: px(0.34), y: py(0.32), w: w * 0.28, h: h * 0.24, fill: { color } });
		ellipse(slide, px(0.14), py(0.62), w * 0.22, h * 0.22, { fill: NO, line: L });
		ellipse(slide, px(0.60), py(0.62), w * 0.22, h * 0.22, { fill: NO, line: L });
	}
}

/** Four small social glyphs (facebook / instagram / twitter / whatsapp). */
function socialIcons(slide, x, y, color) {
	const c = color || '3B3B3B';
	const s = 0.16;
	tx(slide, 'f', { x, y: y - 0.045, w: 0.16, h: 0.24, fontFace: SANS_SB, fontSize: 11, color: c });
	slide.addShape('roundRect', {
		x: x + 0.221, y, w: s, h: s, rectRadius: 0.04,
		fill: { type: 'none' }, line: { color: c, width: 1 },
	});
	ellipse(slide, x + 0.221 + s * 0.3, y + s * 0.3, s * 0.4, s * 0.4, { fill: { type: 'none' }, line: { color: c, width: 0.75 } });
	slide.addShape('custGeom', {
		x: x + 0.502, y: y + 0.02, w: 0.167, h: 0.13, fill: { color: c },
		points: [
			{ x: 0.167, y: 0.0, moveTo: true }, { x: 0.12, y: 0.03 }, { x: 0.06, y: 0.015 },
			{ x: 0.0, y: 0.06 }, { x: 0.045, y: 0.075 }, { x: 0.0, y: 0.12 },
			{ x: 0.09, y: 0.13 }, { x: 0.15, y: 0.07 }, { x: 0.15, y: 0.035 }, { close: true },
		],
	});
	ellipse(slide, x + 0.796, y, s, s, { fill: { color: c } });
	slide.addShape('triangle', { x: x + 0.796, y: y + s * 0.72, w: s * 0.4, h: s * 0.4, fill: { color: c }, rotate: 200 });
}

/* ------------------------------------------------------------------ *
 * Repeated content blocks
 * ------------------------------------------------------------------ */

const LOREM_LONG =
	"There are many variations of passages of Lorem Ipsum available, but the majority have suffered " +
	"alteration in some form, by injected humor, or randomized words which don't look even slightly believable";
const LOREM_SHORT = 'PLACEHOLDER';
const INTRO_LINES = ['Medical is an important factor.', 'Ability to acquire skills ', 'some one.'];
const CARD_BODY = ['But I must explain to you how', 'all this happiness aquan facer.'];
const MISTAKEN =
	'PLACEHOLDER' +
	'mistaken idea of denouncing pleasure and I will give you a complete account of the system, expound master-builder';

/** "Medical is an important factor. / Ability to acquire skills / some one." */
function intro(slide, x, y, color, w) {
	tx(slide, INTRO_LINES.join('\n'), {
		x, y, w: w || 4.364, h: 1.097, fontFace: SANS, fontSize: 18, lineSpacing: 24, color: color || BLACK,
	});
}

/** White "Heart Care" style card with an icon, title and two body lines. */
function infoCard(slide, x, y, title, iconKind, iconColor) {
	roundRect(slide, x, y, 2.844, 2.158, {
		fill: { color: WHITE },
		shadow: { type: 'outer', color: BLACK, opacity: 0.35, blur: 5, offset: 2, angle: 90 },
	});
	icon(slide, iconKind, x + 1.165, y + 0.351, 0.414, 0.367, iconColor);
	tx(slide, title, {
		x: x + 0.222, y: y + 0.954, w: 2.3, h: 0.438,
		fontFace: SANS_M, fontSize: 20, color: BLACK, transparency: 30, align: 'center',
	});
	tx(slide, CARD_BODY.join('\n'), {
		x: x + 0.328, y: y + 1.386, w: 2.087, h: 0.529,
		fontFace: SANS, fontSize: 8.5, color: BLACK, transparency: 30, align: 'center', lineSpacing: 16,
	});
}

/** Big statistic: small "Subtitle" caption, huge number, small footnote. */
function statBlock(slide, x, y, value, note, color) {
	tx(slide, 'Subtitle', { x: x + 0.02, y, w: 0.634, h: 0.296, fontFace: INTER, fontSize: 8, color });
	tx(slide, value, { x, y: y + 0.245, w: 1.4, h: 0.572, fontFace: SANS_M, fontSize: 28, color });
	tx(slide, note, { x: x + 0.02, y: y + 0.745, w: 1.174, h: 0.269, fontFace: SANS_M, fontSize: 10, color });
}

/** Trio of icon bubbles with a caption above each one. */
const TRIO = [
	{ dx: 0, label: '$89%', kind: 'heart', color: BLUE },
	{ dx: 1.295, label: '1.1 Million', kind: 'caduceus', color: GREEN },
	{ dx: 2.59, label: '73 Billion', kind: 'stetho', color: TEAL },
];

/** `inverted` swaps the bubble to white and tints the glyph instead. */
function statTrio(slide, x, y, o) {
	const opt = o || {};
	TRIO.forEach((t) => {
		const cx = x + t.dx;
		ellipse(slide, cx, y, 0.745, 0.745, { fill: { color: opt.inverted ? WHITE : t.color } });
		icon(slide, t.kind, cx + 0.192, y + 0.199, 0.361, 0.348, opt.inverted ? t.color : WHITE);
		tx(slide, t.label, {
			x: cx - 0.214, y: y - 0.59, w: 1.174, h: 0.55,
			fontFace: SANS_M, fontSize: 14, color: opt.labelColor || BLACK, align: 'center', lineSpacing: 36,
		});
	});
}

/** Skill bars: label, grey track, coloured progress, percentage. */
const SKILLS = [
	{ label: 'Photoshop', value: 2.013, pct: '69%' },
	{ label: 'Illustrator', value: 1.37, pct: '61%' },
	{ label: 'UX Design', value: 1.837, pct: '87%' },
];

function skillBars(slide, x, y, colors, font) {
	tx(slide, 'Skill', { x: x - 0.115, y: y - 0.722, w: 0.657, h: 0.37, fontFace: SANS_SB, fontSize: 16 });
	hLine(slide, x - 0.012, y - 0.379, 0.632, BLACK, 1, 70);
	SKILLS.forEach((s, i) => {
		const by = y + i * 0.408;
		tx(slide, s.label, {
			x: x - 0.102, y: by - 0.286, w: 0.952, h: 0.309,
			fontFace: font || SANS, fontSize: 10, color: BLACK, transparency: 30, lineSpacing: 16,
		});
		slide.addShape('roundRect', { x, y: by, w: 2.17, h: 0.073, fill: { color: LT2 }, rectRadius: 0.0365 });
		slide.addShape('roundRect', { x, y: by, w: s.value, h: 0.073, fill: { color: colors[i] }, rectRadius: 0.0365 });
		tx(slide, s.pct, {
			x: x + 2.426, y: by - 0.115, w: 0.567, h: 0.303,
			fontFace: font === INTER ? INTER_SB : SANS_SB, fontSize: 12, color: BLACK, transparency: 12,
		});
	});
}

/** Wingdings-check bullet list used on the "data" slides. */
const CHECKS = ['There are many variations', 'Lorem Ipsum available,', 'Available, but the majority'];

function checkList(slide, x, y, color) {
	CHECKS.forEach((t, i) => {
		tx(slide, t, {
			x, y: y + i * 0.331, w: 2.117, h: 0.303, fontFace: INTER, fontSize: 10, color: color || BLACK,
			bullet: { characterCode: '2713', indent: 13.5 }, lineSpacing: 16,
		});
	});
}

/** Person name + role with a thin rule underneath the role. */
function person(slide, x, y, name, role, o) {
	const opt = o || {};
	tx(slide, name, {
		x, y, w: opt.nameW || 2.789, h: 0.37, fontFace: opt.font || SANS_SB, fontSize: 16,
		color: opt.color || BLACK, align: opt.align || 'left',
	});
	tx(slide, role, {
		x: x + (opt.roleDx || 0), y: y + 0.328, w: opt.roleW || 1.184, h: 0.252,
		fontFace: opt.font === INTER_SB ? INTER : SANS, fontSize: opt.roleSize || 9,
		color: opt.color || BLACK, align: opt.align || 'left',
	});
	hLine(slide, x + (opt.ruleDx || 0.108), y + 0.328, opt.ruleW || 0.91,
		opt.ruleColor || BLACK, 1, opt.ruleFade === undefined ? 70 : opt.ruleFade);
}

/* ------------------------------------------------------------------ *
 * Slides
 * ------------------------------------------------------------------ */

/** Slides 1 and 24 — cover. */
function titleSlide(slide, badgeY) {
	// green wave beneath the picture column (mostly hidden by the aqua one)
	poly(slide, 0, 1.807, 7.428, 5.695, { color: GREEN }, [
		[7.428, 5.695], [0, 5.695], [0, 0.013, 0, 5.695, 0, 0.017],
		[0.907, 0.064, 0, -0.033, 0.857, 0.056],
		[2.665, 0.869, 1.554, 0.172, 2.179, 0.432],
		[3.700, 2.363, 3.124, 1.282, 3.393, 1.842],
		[6.914, 4.719, 4.394, 3.544, 5.416, 4.688],
	]);
	// aqua wave sweeping from the top-left corner down to the bottom edge
	poly(slide, 0, 0, 12.648, 7.502, { color: AQUA }, [
		[0, 0],
		[0.586, 0.217, 0, 0.008, 0.552, 0.201],
		[1.875, 1.086, 1.059, 0.434, 1.495, 0.731],
		[3.120, 2.926, 2.422, 1.597, 2.842, 2.232],
		[3.776, 5.143, 3.408, 3.644, 3.448, 4.440],
		[5.460, 6.456, 4.099, 5.838, 4.714, 6.298],
		[6.980, 6.484, 5.943, 6.559, 6.494, 6.523],
		[9.660, 6.158, 7.880, 6.411, 8.753, 6.131],
		[11.946, 6.855, 10.463, 6.182, 11.287, 6.378],
		[12.321, 7.165, 12.077, 6.950, 12.202, 7.055],
		[12.648, 7.502, 12.373, 7.214, 12.591, 7.502],
		[0, 7.502], [0, 0], 'close',
	]);
	photo(slide, 0, 1.514, 5.853, 5.986);
	ring(slide, 10.612, -1.816, 4.581);
	ring(slide, 10.27, 4.144, 2.712);
	plusCluster(slide, 2.689, 0.718);

	// logo lockup
	ellipse(slide, 4.782, 0.253, 0.544, 0.544, { fill: { color: BLUE } });
	ellipse(slide, 4.873, 0.299, 0.452, 0.452, { fill: { color: WHITE } });
	slide.addShape('mathPlus', { x: 4.956, y: 0.386, w: 0.287, h: 0.257, fill: { color: BLUE } });
	tx(slide, 'COMPANY LOGO', {
		x: 5.376, y: 0.376, w: 1.661, h: 0.269, fontFace: SANS_SB, fontSize: 10, color: BLACK, transparency: 12,
	});

	hLine(slide, 4.782, 1.486, 0.695, BLACK, 7.25, 62);
	tx(slide, 'DESIGN PRESENTATION', {
		x: 4.678, y: 1.616, w: 3.517, h: 0.291, fontFace: SANS, fontSize: 11, charSpacing: 6, lineSpacing: 14,
	});
	tx(slide, 'MEDICAL', { x: 4.607, y: 1.629, w: 7.129, h: 1.868, fontFace: SANS_SB, fontSize: 106, bold: true });
	tx(slide, 'Medical Presentation Design', {
		x: 4.704, y: 3.303, w: 4.525, h: 0.37, fontFace: SANS, fontSize: 16, charSpacing: 3,
	});

	tx(slide, 'Medi By :', { x: 4.705, y: 4.363, w: 1.379, h: 0.269, fontFace: SANS, fontSize: 10 });
	tx(slide, 'CEO/Medical Template', {
		x: 4.716, y: 4.565, w: 2.712, h: 0.325, fontFace: SANS_M, fontSize: 12, lineSpacing: 16,
	});
	tx(slide, 'Presentation Co.', { x: 4.702, y: 4.833, w: 1.556, h: 0.252, fontFace: SANS, fontSize: 9 });
	hLine(slide, 4.819, badgeY + 0.193, 3.04, BLACK, 1, 62);
	pill(slide, 8.06, badgeY);
}

function slide1(s) { titleSlide(s, 4.9); }
function slide24(s) { titleSlide(s, 4.868); }

/** Slide 2 — "Welcome To Our Medical". */
function slide2(s) {
	ring(s, -1.564, -1.809, 4.581);
	// blue disc, clipped flat where it runs off the right edge
	poly(s, 7.791, 1.019, 5.542, 6.333, { color: BLUE }, [
		[3.167, 0],
		[5.406, 0.927, 4.041, 0, 4.833, 0.354],
		[5.542, 1.078], [5.542, 5.256], [5.406, 5.406],
		[3.167, 6.333, 4.833, 5.979, 4.041, 6.333],
		[0, 3.167, 1.418, 6.333, 0, 4.916],
		[3.167, 0, 0, 1.418, 1.418, 0], 'close',
	]);
	photo(s, 7.791, 1.78, 4.81, 4.81, { shape: 'ellipse' });

	tx(s, 'Welcome To Our\nMedical', {
		x: 0.621, y: 1.269, w: 5.146, h: 1.224, fontFace: SANS_SB, fontSize: 36, lineSpacing: 40,
	});
	person(s, 0.726, 3.315, 'Wilkins Micawber', 'CEO/Manager', {
		color: '111A0B', roleDx: -0.014, ruleColor: 'A6A6A6', ruleFade: 70,
	});
	tx(s, LOREM_LONG.replace('alteration in some form, by injected humor, or randomized words which don\'t look even slightly believable',
		'alteration variations of passages of Lorem Ipsum available, but the majority have suffered alteration ' +
		'variations of passages of Lorem Ipsum available, '), {
		x: 0.734, y: 4.04, w: 2.758, h: 1.418, fontFace: INTER, fontSize: 8, lineSpacing: 16,
	});
	pill(s, 0.726, 6.036);
	siteUrl(s, 0.621, 6.796);

	skillBars(s, 4.292, 4.04, [GREEN, BLACK, BLUE]);
}

/** Slide 3 — section header. */
function slide3(s) {
	rect(s, 0, -0.015, 3.254, 7.515, { fill: { color: BLUE } });
	ellipse(s, 0.619, 1.378, 5.206, 5.206, { fill: { type: 'none' }, line: { color: WHITE, width: 1 } });
	photo(s, 0.849, 1.591, 4.81, 4.81, { shape: 'ellipse' });
	ring(s, 10.901, 4.294, 4.581);
	plusCluster(s, 3.538, 0.502);

	tx(s, 'Building A\nHealthier Future,\nTogether', {
		x: 6.382, y: 1.494, w: 6.546, h: 2.471, fontFace: SANS_SB, fontSize: 50, lineSpacing: 56,
	});
	tx(s, LOREM_LONG, { x: 6.48, y: 4.459, w: 4.298, h: 0.976, fontFace: INTER, fontSize: 10, lineSpacing: 16 });
	pill(s, 6.605, 5.671);
	siteUrl(s, 6.474, 6.796);
}

/** Slide 4 — "Leading The Way In Medical Excellence". */
function slide4(s) {
	ring(s, -1.564, -2.094, 4.581);
	photo(s, 9.238, 2.828, 4.096, 4.672);
	tx(s, 'Leading The Way In\nMedical Excellence', {
		x: 0.559, y: 0.771, w: 5.146, h: 1.347, fontFace: SANS_SB, fontSize: 36, lineSpacing: 45,
	});

	// two icon bubbles, each with two paragraphs of copy
	[
		{ bx: 0.609, by: 2.828, tx0: 1.526, ty: 2.912, color: BLUE, kind: 'heart' },
		{ bx: 4.838, by: 2.850, tx0: 5.748, ty: 2.904, color: GREEN, kind: 'stetho' },
	].forEach((c) => {
		ellipse(s, c.bx, c.by, 0.726, 0.726, { fill: { color: c.color } });
		icon(s, c.kind, c.bx + 0.181, c.by + 0.208, 0.366, 0.334, WHITE);
		tx(s, 'Example Text', { x: c.tx0, y: c.ty, w: 2.071, h: 0.286, fontFace: INTER_M, fontSize: 11 });
		tx(s, LOREM_SHORT, {
			x: c.tx0, y: c.ty + 0.301, w: 2.688, h: 0.527, fontFace: INTER, fontSize: 8, lineSpacing: 16,
		});
		tx(s, LOREM_SHORT, {
			x: c.tx0 - 0.005, y: c.ty + 0.983, w: 2.688, h: 0.527, fontFace: INTER, fontSize: 8, lineSpacing: 16,
		});
	});

	infoCard(s, 1.521, 4.833, 'Heart Care', 'heart', BLUE);
	infoCard(s, 5.83, 4.84, 'Dental Care', 'tooth', GREEN);
}

/** Slide 5 — two testimonial cards on a green band. */
function slide5(s) {
	photo(s, 0, 2.344, 13.333, 5.156);
	rect(s, -0.006, 2.344, 13.339, 5.156, { fill: { color: GREEN, transparency: 54 } });

	tx(s, 'Making Health Care Better Togethers', {
		x: 0.705, y: 0.664, w: 5.513, h: 1.224, fontFace: SANS_SB, fontSize: 36, lineSpacing: 40,
	});
	intro(s, 6.9, 0.727);

	[
		{ x: 0.796, name: 'Alon Walker', role: 'CEO/Manager', nx: 1.302, nw: 1.617, rx: 1.518 },
		{ x: 6.884, name: 'Robert William', role: 'Photographer', nx: 6.533, nw: 2.038, rx: 6.923 },
	].forEach((c) => {
		roundRect(s, c.x, 3.052, 5.654, 3.146, {
			fill: { color: WHITE },
			shadow: { type: 'outer', color: BLACK, opacity: 0.35, blur: 5, offset: 2, angle: 90 },
		});
		photo(s, c.x + 0.427, 3.408, 1.714, 1.79, { shape: 'roundRect' });
		tx(s, 'Description Title Goes Here', {
			x: c.x + 2.452, y: 3.307, w: 2.071, h: 0.572, fontFace: SANS_SB, fontSize: 14, color: BLACK, transparency: 28,
		});
		tx(s, MISTAKEN, {
			x: c.x + 2.452, y: 3.884, w: 2.969, h: 1.592, fontFace: SANS, fontSize: 10, color: BLACK,
			transparency: 30, lineSpacing: 18,
		});
		tx(s, c.name, {
			x: c.nx, y: 5.352, w: c.nw, h: 0.37, fontFace: SANS_SB, fontSize: 16, color: BLACK,
			transparency: 30, align: 'center',
		});
		tx(s, c.role, {
			x: c.rx, y: 5.68, w: 1.184, h: 0.269, fontFace: SANS, fontSize: 10, color: BLACK,
			transparency: 30, align: 'center',
		});
		hLine(s, c.nx + 0.161, 5.637, 1.294, BLACK, 1, 70);
		for (let i = 0; i < 5; i++) {
			// the original glyph is a hollow outlined star
			s.addShape('star5', {
				x: c.x + 2.55 + i * 0.303, y: 5.557, w: 0.194, h: 0.185,
				fill: { type: 'none' }, line: { color: GREEN, width: 0.75 },
			});
		}
	});
}

/** Slide 6 — "Bringing Health To Life The Whole Family". */
function slide6(s) {
	ring(s, 4.926, 2.307, 4.581);
	rect(s, 7.973, 1.501, 5.36, 5.999, { fill: { color: GREEN } });
	photo(s, 9.831, 2.234, 3.166, 2.334);
	photo(s, 9.831, 4.788, 3.166, 2.334);

	tx(s, 'Bringing Health To Life \nThe Whole Family', {
		x: 0.679, y: 0.936, w: 6.515, h: 1.347, fontFace: SANS_SB, fontSize: 36, lineSpacing: 45,
	});
	tx(s, 'There are many variations\nof passages', {
		x: 0.751, y: 2.965, w: 3.2, h: 0.55, fontFace: SANS, fontSize: 14, color: BLACK, transparency: 12, lineSpacing: 16,
	});
	tx(s, 'There are many variations of passages of Lorem Ipsum available, but the majority alteration in some ' +
		'form, by injected humor, or randomized are many variations of passages of Lorem Ipsum available, but the majority', {
		x: 0.728, y: 3.724, w: 3.574, h: 1.2, fontFace: INTER, fontSize: 10, lineSpacing: 16,
	});
	plusCluster(s, 3.84, 5.513);
	pill(s, 0.795, 6.036);
	siteUrl(s, 0.664, 6.796);

	infoCard(s, 6.621, 2.234, 'Heart Care', 'heart', BLUE);
	infoCard(s, 6.621, 4.803, 'Dental Care', 'tooth', GREEN);
}

/** Slide 7 — "The Future of Healthcare". */
function slide7(s) {
	ring(s, 10.612, -1.816, 4.581);
	rect(s, 0, 1.958, 5.478, 2.854, { fill: { color: BLUE } });
	photo(s, 0.663, 3.75, 2.749, 3.348);
	photo(s, 3.75, 3.75, 3.211, 3.348);

	tx(s, 'The Future of Healthcare', {
		x: 0.559, y: 0.632, w: 6.785, h: 0.695, fontFace: INTER_SB, fontSize: 36, lineSpacing: 45,
	});
	intro(s, 0.663, 2.433, WHITE, 4.226);

	[
		{ y: 2.502, color: BLUE, kind: 'caduceus', ty: 2.427, by: 2.684 },
		{ y: 4.264, color: AQUA, kind: 'stetho', ty: 4.189, by: 4.45 },
		{ y: 6.037, color: TEAL, kind: 'heart', ty: 5.962, by: 6.223 },
	].forEach((r) => {
		rect(s, 6.416, r.y, 1.006, 1.061, { fill: { color: r.color } });
		icon(s, r.kind, 6.638, r.y + 0.246, 0.569, 0.569, WHITE);
		tx(s, 'Example Text', { x: 7.844, y: r.ty, w: 3.263, h: 0.451, fontFace: INTER_M, fontSize: 11 });
		tx(s, LOREM_SHORT + 'PLACEHOLDER', {
			x: 7.844, y: r.by, w: 3.887, h: 0.745, fontFace: INTER, fontSize: 8, lineSpacing: 16,
		});
	});
}

/** Slide 8 — "We Care For Your Health" with numbered list. */
function slide8(s) {
	ring(s, 10.612, -1.816, 4.581);
	rect(s, -0.014, 5.772, 4.937, 1.728, { fill: { color: BLUE } });

	tx(s, 'We Care For Your Health', {
		x: 0.573, y: 0.632, w: 6.588, h: 0.695, fontFace: INTER_SB, fontSize: 36, lineSpacing: 45,
	});
	intro(s, 0.559, 1.344);
	plusCluster(s, 1.742, 3.085);
	infoCard(s, 3.749, 2.697, 'Heart Care', 'heart', BLUE);
	infoCard(s, 0.592, 4.83, 'Dental Care', 'tooth', GREEN);

	// icon tiles + head-line statistic
	[
		{ x: 5.628, y: 5.235, kind: 'stetho', color: BLUE },
		{ x: 6.51, y: 5.235, kind: 'caduceus', color: GREEN },
		{ x: 5.628, y: 6.173, kind: 'heart', color: TEAL },
	].forEach((t) => {
		rect(s, t.x, t.y, 0.766, 0.808, { fill: { type: 'none' }, line: { color: 'A6A6A6', width: 0.75 } });
		icon(s, t.kind, t.x + 0.166, t.y + 0.19, 0.433, 0.433, t.color);
	});
	tx(s, 'Subtitle', { x: 6.452, y: 6.045, w: 0.634, h: 0.296, fontFace: INTER, fontSize: 8 });
	tx(s, '150M', { x: 6.432, y: 6.29, w: 1.354, h: 0.572, fontFace: SANS_M, fontSize: 28 });
	tx(s, '1.1 Million', { x: 6.452, y: 6.79, w: 1.174, h: 0.269, fontFace: SANS_M, fontSize: 10 });

	['01', '02', '03'].forEach((n, i) => {
		const y = [3.037, 4.439, 5.896][i];
		ellipse(s, 8.45, y + 0.202, 0.486, 0.486, { fill: { color: [BLUE, GREEN, TEAL][i] } });
		tx(s, n, { x: 8.68, y, w: 0.777, h: 0.572, fontFace: INTER_T, fontSize: 28 });
		tx(s, 'Example Text', { x: 9.476, y: y + 0.039, w: 2.071, h: 0.286, fontFace: INTER_M, fontSize: 11 });
		tx(s, LOREM_SHORT, { x: 9.476, y: y + 0.374, w: 2.383, h: 0.527, fontFace: INTER, fontSize: 8, lineSpacing: 16 });
	});
}

/** Slide 9 — "Get Quick Medical Services". */
function slide9(s) {
	ring(s, 2.26, 3.929, 4.581);
	rect(s, 7.872, 2.983, 5.461, 4.517, { fill: { color: BLUE } });
	photo(s, 0.734, 2.983, 3.425, 2.557);
	photo(s, 4.447, 2.983, 3.425, 2.557);

	tx(s, 'Get Quick Medical Services', {
		x: 0.679, y: 0.698, w: 5.987, h: 1.347, fontFace: SANS_SB, fontSize: 36, lineSpacing: 45,
	});
	intro(s, 6.996, 0.808);
	statBlock(s, 0.605, 5.782, '126K', '1.1 Million', BLACK);
	statBlock(s, 4.339, 5.782, '173M', '43.7 Billion', BLACK);

	[
		{ y: 3.322, kind: 'heart', ty: 3.304, by: 3.694 },
		{ y: 4.789, kind: 'stetho', ty: 4.722, by: 5.133 },
		{ y: 6.124, kind: 'caduceus', ty: 6.126, by: 6.502 },
	].forEach((r) => {
		ellipse(s, 8.78, r.y, 0.716, 0.716, { fill: { color: WHITE } });
		icon(s, r.kind, 8.96, r.y + 0.19, 0.361, 0.348, BLUE);
		tx(s, 'Example Text', { x: 9.71, y: r.ty, w: 2.071, h: 0.286, fontFace: INTER_M, fontSize: 11, color: WHITE });
		tx(s, 'There are many variations of passages Ipsum but the majority', {
			x: 9.71, y: r.by, w: 2.631, h: 0.527, fontFace: INTER, fontSize: 8, color: WHITE, lineSpacing: 16,
		});
	});
}

/** Slide 10 — "Healthcare With Hearth, Foe Every Concern." */
function slide10(s) {
	ring(s, 9.932, -2.163, 4.581);
	rect(s, 8.16, 3.032, 5.173, 4.468, { fill: { color: GREEN } });
	[0.739, 3.301, 5.854].forEach((x) => photo(s, x, 3.032, 2.306, 2.254));

	tx(s, 'Healthcare With\nHearth, Foe Every\nConcern.', {
		x: 0.651, y: 0.74, w: 5.146, h: 1.795, fontFace: SANS_SB, fontSize: 36, lineSpacing: 40,
	});
	intro(s, 6.043, 0.743);

	[
		{ x: 0.63, y: 5.671, value: '50K', w: 0.932 },
		{ x: 3.257, y: 5.678, value: '73%', w: 0.932 },
		{ x: 5.868, y: 5.671, value: '1.4M', w: 1.261 },
	].forEach((c) => {
		tx(s, c.value, { x: c.x, y: c.y, w: c.w, h: 0.55, fontFace: INTER_SB, fontSize: 24, lineSpacing: 35 });
		tx(s, 'Your Title Here', { x: c.x, y: c.y + 0.532, w: 2.071, h: 0.286, fontFace: INTER_M, fontSize: 11 });
		tx(s, 'There are many variations of passages majority many', {
			x: c.x, y: c.y + 0.904, w: 2.316, h: 0.527, fontFace: INTER, fontSize: 8, lineSpacing: 16,
		});
	});

	tx(s, 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered', {
		x: 8.874, y: 3.179, w: 4.107, h: 0.527, fontFace: INTER, fontSize: 10, color: WHITE, lineSpacing: 16,
	});
	checkList(s, 8.874, 3.989, WHITE);
	statTrio(s, 9.089, 6.017, { inverted: true, labelColor: WHITE });
}

/** Slide 11 — "The Best Choice For Our Healthy". */
function slide11(s) {
	ring(s, -1.564, -2.094, 4.581);
	rect(s, 5.609, 3.75, 7.724, 3.75, { fill: { color: BLUE } });
	tx(s, 'The Best Choice For Our Healthy', {
		x: 5.515, y: 0.91, w: 6.16, h: 1.224, fontFace: SANS_SB, fontSize: 36, lineSpacing: 40,
	});
	plusCluster(s, 0.579, 4.196);
	infoCard(s, 1.836, 2.434, 'Heart Care', 'heart', BLUE);
	infoCard(s, 1.836, 5.007, 'Dental Care', 'tooth', GREEN);

	[
		{ x: 5.609, kind: 'stetho' }, { x: 8.082, kind: 'caduceus' }, { x: 10.459, kind: 'heart' },
	].forEach((c, i) => {
		photo(s, c.x, i === 1 ? 2.742 : 2.723, 1.764, 1.947);
		icon(s, c.kind, c.x + 0.598, 5.028, 0.569, 0.569, WHITE);
		tx(s, 'Example Text', {
			x: c.x + 0.063, y: 5.843, w: 1.639, h: 0.286, fontFace: INTER_M, fontSize: 11, color: WHITE, align: 'center',
		});
		tx(s, LOREM_SHORT, {
			x: c.x - 0.094, y: 6.11, w: 1.954, h: 0.752, fontFace: INTER, fontSize: 8, color: WHITE,
			align: 'center', lineSpacing: 16,
		});
	});
}

/** Slide 12 — "Your Health is Our Goals". */
function slide12(s) {
	rect(s, 0, 0, 5.285, 7.5, { fill: { color: GREEN } });
	ring(s, 10.612, -2.467, 4.581);
	tx(s, 'Your Health is \nOur Goals', {
		x: 0.555, y: 0.632, w: 3.837, h: 1.326, fontFace: INTER_SB, fontSize: 36, color: WHITE, lineSpacing: 45,
	});
	tx(s, 'There are many variations\nof passages', {
		x: 0.56, y: 2.6, w: 3.2, h: 0.55, fontFace: SANS, fontSize: 14, color: WHITE, lineSpacing: 16,
	});
	statBlock(s, 0.557, 3.766, '126K', '1.1 Million', WHITE);
	statBlock(s, 3.038, 3.766, '173M', '43.7 Billion', WHITE);
	photo(s, 0.673, 5.142, 2.124, 2.368);
	photo(s, 3.161, 5.142, 2.124, 2.368);

	tx(s, 'PLACEHOLDER' +
		'complete account of the system, expound master-builder of human happiness', {
		x: 6.707, y: 1.887, w: 5.308, h: 0.834, fontFace: SANS, fontSize: 10, color: BLACK, transparency: 30, lineSpacing: 18,
	});
	[
		{ y: 3.412, name: 'Alon Walker', dot: BLUE }, { y: 5.292, name: 'Alaska Ethan', dot: AQUA },
	].forEach((p) => {
		ellipse(s, 6.071, p.y + 0.097, 0.358, 0.358, { fill: { color: p.dot } });
		tx(s, p.name, { x: 6.713, y: p.y, w: 1.733, h: 0.337, fontFace: SANS_SB, fontSize: 14 });
		tx(s, 'Position Title Here', { x: 6.713, y: p.y + 0.23, w: 1.449, h: 0.309, fontFace: SANS, fontSize: 10, lineSpacing: 16 });
		tx(s, 'PLACEHOLDER', {
			x: 6.717, y: p.y + 0.583, w: 2.223, h: 0.752, fontFace: SANS, fontSize: 10, lineSpacing: 16,
		});
	});

	// gradient "Description" panel (green corner fading into blue)
	roundRect(s, 10.167, 3.153, 1.993, 3.683, { fill: { color: BLUE } });
	cornerFade(s, 10.167, 3.153, 1.993, 3.683, GREEN, BLUE);
	tx(s, 'Description', {
		x: 10.312, y: 3.535, w: 1.386, h: 0.337, fontFace: SANS_SB, fontSize: 14, color: WHITE, align: 'center',
	});
	tx(s, 'But I must denouncing the master-builder of human happiness', {
		x: 10.312, y: 3.812, w: 1.703, h: 0.829, fontFace: SANS, fontSize: 8.5, color: WHITE, align: 'center', lineSpacing: 18,
	});
	hLine(s, 10.707, 4.812, 0.912, WHITE, 1);
	tx(s, 'But I must denouncing the master-builder of human happiness', {
		x: 10.312, y: 4.984, w: 1.703, h: 0.829, fontFace: SANS, fontSize: 8.5, color: WHITE, align: 'center', lineSpacing: 18,
	});
	tx(s, '30+B', {
		x: 10.592, y: 5.984, w: 1.143, h: 0.563, fontFace: SANS_SB, fontSize: 24, color: WHITE, align: 'center', lineSpacing: 35,
	});
}

/**
 * Diagonal green -> blue gradient (the original uses a 300° linear gradFill,
 * which pptxgenjs cannot express) approximated with chamfered corner bands
 * layered over an already-blue card.
 */
function cornerFade(slide, x, y, w, h, from, to) {
	const STEPS = 14;
	const span = (w + h) * 0.55; // distance from the top-left corner where `to` takes over
	for (let i = STEPS; i >= 1; i--) {
		const d = (span * i) / STEPS;
		poly(slide, x, y, w, h, { color: mix(from, to, (i - 0.5) / STEPS) }, [
			[Math.min(d, w), 0], [CARD_R, 0], [0, CARD_R], [0, Math.min(d, h)], 'close',
		]);
	}
}

/** Slide 13 — "Our Leader". */
function slide13(s) {
	ring(s, -1.962, 1.378, 4.581);
	// blue slab with a rounded left edge
	poly(s, 7.333, 1.976, 6.0, 4.036, { color: BLUE }, [
		[2.018, 0], [5.935, 0], [6.0, 0.003], [6.0, 4.032], [5.935, 4.036], [2.018, 4.036],
		[0.010, 2.224, 0.973, 4.036, 0.114, 3.242],
		[0, 2.018], [0.010, 1.811],
		[2.018, 0, 0.114, 0.794, 0.973, 0], 'close',
	]);
	photo(s, 7.333, 1.952, 4.06, 4.06, { shape: 'ellipse' });

	tx(s, 'Our Leader', { x: 3.413, y: 0.695, w: 3.269, h: 0.827, fontFace: INTER_SB, fontSize: 36, lineSpacing: 55 });
	person(s, 3.413, 2.771, 'Alaska Ethan', 'CEO/Manager', {
		font: INTER_SB, nameW: 1.69, roleSize: 10, ruleDx: 0.099, ruleW: 1.39,
	});
	tx(s, LOREM_LONG, { x: 3.415, y: 3.52, w: 3.3, h: 1.2, fontFace: INTER, fontSize: 10, lineSpacing: 16 });
	socialIcons(s, 3.512, 4.888);
	plusCluster(s, 6.582, 5.485);
	pill(s, 0.795, 6.242);
	siteUrl(s, 0.664, 6.796);
}

/** Slide 14 — "Our Expert Team". */
function slide14(s) {
	rect(s, -0.014, 0, 5.172, 3.245, { fill: { color: BLUE } });
	ring(s, 10.612, -1.816, 4.581);
	tx(s, 'Our Expert Team', {
		x: 0.545, y: 1.325, w: 4.326, h: 0.8, fontFace: INTER_SB, fontSize: 36, color: WHITE, lineSpacing: 55,
	});
	skillBars(s, 2.128, 4.254, [BLUE, BLUE, BLACK], INTER);
	pill(s, 0.795, 6.036);
	siteUrl(s, 0.664, 6.796);
	intro(s, 5.677, 1.855);

	[
		{ x: 5.731, name: 'Alon Walker' }, { x: 8.091, name: 'Alaska Ethan' }, { x: 10.451, name: 'Robert Jonson' },
	].forEach((c) => {
		photo(s, c.x, 3.251, 2.081, 2.3);
		tx(s, c.name, { x: c.x - 0.1, y: 5.694, w: 1.788, h: 0.337, fontFace: INTER_SB, fontSize: 14 });
		tx(s, 'Position Title Here', {
			x: c.x - 0.1, y: 5.924, w: 1.449, h: 0.303, fontFace: INTER, fontSize: 10, lineSpacing: 16,
		});
		socialIcons(s, c.x, 6.304);
	});
}

/** Slide 15 — "Our Hospital Medical Photography". */
function slide15(s) {
	ring(s, 10.612, -2.467, 4.581);
	photo(s, 0, 2.972, 5.389, 4.528);
	tx(s, 'Our Hospital Medical Photography', {
		x: 0.616, y: 0.674, w: 6.383, h: 1.347, fontFace: SANS_SB, fontSize: 36, lineSpacing: 45,
	});
	tx(s, 'Medical is an important factor. Ability to acquire skills some one.', {
		x: 5.948, y: 2.044, w: 5.715, h: 0.761, fontFace: SANS, fontSize: 18, lineSpacing: 24,
	});

	[
		{ x: 5.948, y: 3.624, n: '01', color: BLUE }, { x: 9.64, y: 3.618, n: '02', color: GREEN },
		{ x: 5.948, y: 5.358, n: '03', color: TEAL }, { x: 9.643, y: 5.368, n: '04', color: AQUA },
	].forEach((c) => {
		ellipse(s, c.x, c.y, 0.595, 0.595, { fill: { color: c.color } });
		tx(s, c.n, { x: c.x + 0.243, y: c.y - 0.155, w: 0.855, h: 0.64, fontFace: INTER_T, fontSize: 32 });
		tx(s, 'Your Title Here', { x: c.x + 0.904, y: c.y - 0.121, w: 1.449, h: 0.286, fontFace: INTER_M, fontSize: 11 });
		tx(s, 'Position Title Here', {
			x: c.x + 0.904, y: c.y + 0.109, w: 1.449, h: 0.303, fontFace: INTER, fontSize: 8, lineSpacing: 16,
		});
		tx(s, 'There are many of variations passages Ipsum but the', {
			x: c.x + 0.908, y: c.y + 0.392, w: 2.366, h: 0.527, fontFace: INTER, fontSize: 8, lineSpacing: 16,
		});
	});
}

/** Slide 16 — three description cards, the middle one highlighted. */
function slide16(s) {
	ring(s, 9.932, -2.163, 4.581);
	tx(s, 'We Care For Your Health', {
		x: 0.679, y: 0.634, w: 4.779, h: 1.347, fontFace: SANS_SB, fontSize: 36, lineSpacing: 45,
	});
	intro(s, 6.043, 0.743);
	plusCluster(s, 4.866, 1.702);

	[
		{ x: 0.79, kind: 'medkit', bubble: GREEN, glyph: WHITE, text: BLACK, highlight: false },
		{ x: 4.817, kind: 'heart', bubble: WHITE, glyph: BLUE, text: WHITE, highlight: true },
		{ x: 8.979, kind: 'ambulance', bubble: BLUE, glyph: WHITE, text: BLACK, highlight: false },
	].forEach((c) => {
		if (c.highlight) {
			roundRect(s, c.x, 3.397, 3.564, 3.46, { fill: { color: BLUE } });
			cornerFade(s, c.x, 3.397, 3.564, 3.46, GREEN, BLUE);
		} else {
			roundRect(s, c.x, 3.397, 3.564, 3.46, {
				fill: { color: WHITE },
				shadow: { type: 'outer', color: BLACK, opacity: 0.3, blur: 6, offset: 2, angle: 90 },
			});
		}
		ellipse(s, c.x + 0.333, 3.791, 0.716, 0.716, { fill: { color: c.bubble } });
		icon(s, c.kind, c.x + 0.552, 4.010, 0.278, 0.278, c.glyph);
		tx(s, 'Description Title\nGoes Here', {
			x: c.x + 0.241, y: 4.81, w: 2.071, h: 0.572, fontFace: SANS_SB, fontSize: 14,
			color: c.text, transparency: c.highlight ? 0 : 28,
		});
		tx(s, 'But I must explain to you how mistaken idea of denouncing pleasure I will give master-builder happiness', {
			x: c.x + 0.241, y: 5.446, w: 3.198, h: 0.834, fontFace: SANS, fontSize: 10,
			color: c.text, transparency: c.highlight ? 0 : 30, lineSpacing: 18,
		});
	});
}

/** Slide 17 — "Making Health Care Better Togethers". */
function slide17(s) {
	ring(s, 9.932, -2.163, 4.581);
	photo(s, 0.675, 2.417, 3.188, 5.083);
	photo(s, 3.993, 2.417, 2.459, 2.821);
	tx(s, 'Making Health Care Better Togethers', {
		x: 0.565, y: 0.784, w: 7.276, h: 1.224, fontFace: SANS_SB, fontSize: 36, lineSpacing: 40,
	});
	intro(s, 6.881, 2.497);
	plusCluster(s, 10.975, 3.249);

	roundRect(s, 2.489, 4.484, 2.844, 2.158, {
		fill: { color: WHITE },
		shadow: { type: 'outer', color: BLACK, opacity: 0.35, blur: 5, offset: 2, angle: 90 },
	});
	icon(s, 'heart', 3.654, 4.834, 0.414, 0.367, BLUE);
	tx(s, 'Health Care', {
		x: 2.711, y: 5.438, w: 2.3, h: 0.438, fontFace: SANS_M, fontSize: 20, color: BLACK, transparency: 30, align: 'center',
	});
	tx(s, CARD_BODY.join('\n'), {
		x: 2.817, y: 5.87, w: 2.087, h: 0.529, fontFace: SANS, fontSize: 8.5, color: BLACK,
		transparency: 30, align: 'center', lineSpacing: 16,
	});

	statTrio(s, 6.881, 4.907);
	[6.732, 8.899, 11.066].forEach((x) => {
		tx(s, 'Example Text', { x, y: 6.055, w: 1.251, h: 0.286, fontFace: INTER_M, fontSize: 11 });
		tx(s, 'There are many variations of passages Ipsum', {
			x: x + 0.004, y: 6.341, w: 2.071, h: 0.527, fontFace: INTER, fontSize: 8, lineSpacing: 16,
		});
	});
}

/** Slide 18 — "Checkup Your Health At Our Hospital". */
function slide18(s) {
	ring(s, 10.612, -2.467, 4.581);
	photo(s, 5.048, 2.647, 3.638, 4.853);
	tx(s, 'Checkup Your  Health At Our Hospital', {
		x: 0.545, y: 0.426, w: 5.746, h: 1.326, fontFace: INTER_SB, fontSize: 36, lineSpacing: 45,
	});
	statBlock(s, 0.557, 2.341, '126K', '1.1 Million', BLACK);
	tx(s, 'It is a long established fact that a\nReader will be distracted by the\nReadable of Medical', {
		x: 0.557, y: 3.559, w: 2.836, h: 0.682, fontFace: SANS, fontSize: 10.5, lineSpacing: 14,
	});
	plusCluster(s, 3.554, 4.522);
	pill(s, 0.795, 6.036);
	siteUrl(s, 0.664, 6.796);

	[
		{ y: 2.667, kind: 'heart', color: BLUE }, { y: 4.032, kind: 'stetho', color: GREEN },
		{ y: 5.417, kind: 'medkit', color: TEAL },
	].forEach((r) => {
		icon(s, r.kind, 9.455, r.y + 0.093, 0.433, 0.42, r.color);
		tx(s, 'Example Text', { x: 10.241, y: r.y, w: 2.071, h: 0.337, fontFace: INTER_SB, fontSize: 14 });
		tx(s, 'There are many variations of passages Ipsum', {
			x: 10.241, y: r.y + 0.423, w: 2.356, h: 0.527, fontFace: INTER, fontSize: 10, lineSpacing: 16,
		});
	});
}

/** Slide 19 — "Confidence For The Good Life". */
function slide19(s) {
	ring(s, 10.612, -2.467, 4.581);
	photo(s, 0, 0, 4.098, 7.502);
	photo(s, 4.294, 2.732, 3.579, 4.768);
	tx(s, 'Confidence For The Good Life', {
		x: 6.694, y: 0.886, w: 5.0, h: 1.326, fontFace: INTER_SB, fontSize: 36, lineSpacing: 45,
	});
	intro(s, 8.254, 2.612);
	tx(s, LOREM_LONG, { x: 8.286, y: 3.955, w: 4.622, h: 0.976, fontFace: INTER, fontSize: 10, lineSpacing: 16 });
	plusCluster(s, 5.316, 0.727);

	roundRect(s, 2.872, 1.181, 2.844, 2.158, {
		fill: { color: WHITE },
		shadow: { type: 'outer', color: BLACK, opacity: 0.35, blur: 5, offset: 2, angle: 90 },
	});
	icon(s, 'heart', 4.037, 1.531, 0.414, 0.367, BLUE);
	tx(s, 'Health Care', {
		x: 3.094, y: 2.135, w: 2.3, h: 0.438, fontFace: SANS_M, fontSize: 20, color: BLACK, transparency: 30, align: 'center',
	});
	tx(s, CARD_BODY.join('\n'), {
		x: 3.2, y: 2.567, w: 2.087, h: 0.529, fontFace: SANS, fontSize: 8.5, color: BLACK,
		transparency: 30, align: 'center', lineSpacing: 16,
	});

	statTrio(s, 8.348, 6.049);
}

/** Slide 20 — "Bringing Health to Life the Whole Family". */
function slide20(s) {
	ring(s, -1.246, -2.403, 4.581);
	photo(s, 5.323, 2.681, 3.544, 4.819);
	photo(s, 9.282, 2.681, 3.544, 4.819);
	tx(s, 'Bringing Health to Life the Whole Family', {
		x: 5.269, y: 0.743, w: 7.132, h: 1.326, fontFace: INTER_SB, fontSize: 36, lineSpacing: 45,
	});
	tx(s, 'There are many variations\nof passages', {
		x: 0.574, y: 2.448, w: 3.2, h: 0.55, fontFace: SANS, fontSize: 14, lineSpacing: 16,
	});
	tx(s, 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered', {
		x: 0.588, y: 3.449, w: 4.107, h: 0.527, fontFace: INTER, fontSize: 10, lineSpacing: 16,
	});
	checkList(s, 0.588, 4.259);
	plusCluster(s, 3.887, 2.07);
	statTrio(s, 0.577, 6.084);
}

/** Slide 21 — clustered bar chart. */
function slide21(s) {
	ring(s, 10.612, -2.467, 4.581);
	tx(s, 'Data Chart', { x: 0.558, y: 0.626, w: 3.95, h: 0.707, fontFace: INTER_SB, fontSize: 36 });
	tx(s, 'Medical is an important factor. Ability to acquire skills some one.', {
		x: 0.559, y: 1.528, w: 5.715, h: 0.761, fontFace: SANS, fontSize: 18, lineSpacing: 24,
	});
	tx(s, 'There are many variations\nof passages', {
		x: 0.574, y: 2.797, w: 3.2, h: 0.55, fontFace: SANS, fontSize: 14, lineSpacing: 16,
	});
	tx(s, 'There are many variations of passages of Lorem Ipsum available, but the majority have suffered', {
		x: 0.588, y: 3.528, w: 4.107, h: 0.527, fontFace: INTER, fontSize: 10, lineSpacing: 16,
	});
	checkList(s, 0.588, 4.339);
	pill(s, 0.795, 6.036);
	siteUrl(s, 0.664, 6.796);

	const labels = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
	s.addChart('bar', [
		{ name: 'Series 1', labels, values: [4.3, 2.5, 3.5, 4.5] },
		{ name: 'Series 2', labels, values: [2.4, 4.4, 1.8, 2.8] },
		{ name: 'Series 3', labels, values: [2, 2, 3, 5] },
	], {
		x: 5.778, y: 2.375, w: 6.903, h: 4.602,
		barDir: 'bar', barGapWidthPct: 182, plotArea: { fill: { color: WHITE } },
		chartColors: [AQUA, TEAL, GREEN],
		showLegend: true, legendPos: 'b', legendFontFace: INTER, legendFontSize: 9, legendColor: '595959',
		catAxisLineShow: true, catAxisLabelFontFace: INTER, catAxisLabelFontSize: 10, catAxisLabelColor: '595959',
		valAxisLineShow: false, valAxisLabelFontFace: INTER, valAxisLabelFontSize: 10, valAxisLabelColor: '595959',
		valGridLine: { style: 'solid', color: 'D9D9D9', size: 0.75 },
		catGridLine: { style: 'none' },
		border: { pt: 0, color: 'FFFFFF' }, fill: 'FFFFFF',
	});
}

/** Slide 22 — "Our Device" (desktop monitor mock-up). */
function slide22(s) {
	ring(s, -1.162, 4.801, 4.581);
	tx(s, 'Our Device', { x: 0.558, y: 0.626, w: 3.95, h: 0.707, fontFace: INTER_SB, fontSize: 36 });
	intro(s, 0.6, 1.814);
	tx(s, LOREM_LONG, { x: 0.632, y: 3.157, w: 4.622, h: 0.976, fontFace: INTER, fontSize: 10, lineSpacing: 16 });
	plusCluster(s, 4.776, 5.339);

	// monitor: bezel, screen, neck and base
	s.addShape('roundRect', { x: 6.25, y: 1.373, w: 6.002, h: 3.708, rectRadius: 0.08, fill: { color: '32373C' } });
	ellipse(s, 9.222, 1.47, 0.06, 0.06, { fill: { color: '557DB1' } });
	photo(s, 6.442, 1.559, 5.6, 3.242);
	s.addShape('roundRect', { x: 6.25, y: 5.081, w: 6.002, h: 0.66, rectRadius: 0.06, fill: { color: 'E7E7E8' } });
	s.addShape('trapezoid', { x: 8.451, y: 5.741, w: 1.6, h: 0.63, fill: { color: 'CFD2D5' } });
	s.addShape('roundRect', { x: 7.851, y: 6.29, w: 2.8, h: 0.22, rectRadius: 0.1, fill: { color: 'E7E7E8' } });
}

/**
 * Coarse silhouette of the world map on slide 23, one character per cell.
 * '#' = land.  Rendered as horizontal runs tinted green (west) to teal (east).
 */
const WORLD = [
	'.....................###.....###....................................................',
	'.................##.###.###########.................................................',
	'....................#..###########........##...................#....................',
	'............#......#.....#########............................###...................',
	'..........##.....#........########..................#......######...................',
	'............###....###.....#######.....................#.##################.........',
	'.############...##.#..##...######..........#####.....###..#######################...',
	'##############..####..##...###....#.......##.##..############..##################.#.',
	'.#################..........#............##.#################..#################....',
	'..#....##########....##..................##...###########################...........',
	'........###########..####.............#.....###########################.....#.......',
	'.........###########.#####...........#..######################.#########............',
	'..........##############...............#################################............',
	'..........#############................###.############################.............',
	'..........############...............##.....##########################..............',
	'...........##########...........................###################..#.#............',
	'............########.................#####..#..#####################................',
	'.............###....................###########.###.################................',
	'..............##...................#########..#######..###########..................',
	'...............##.#................#########..##.###....###..###....................',
	'..................##...............###############.......#....###...................',
	'.....................####...........###############.................................',
	'.....................######..............#########................#.................',
	'.....................###..##.............########................##.................',
	'....................#####.#####...........######.......................##...........',
	'.....................##########............######...................................',
	'......................########............#######.#..................#..#...........',
	'.......................#######............#####...#................####.##..........',
	'.......................#####...............####...................#########.........',
	'.......................#####...............####...................#########.........',
	'.......................####.................#.....................#....####.........',
	'......................####..............................................##..........',
	'......................##............................................................',
	'......................##............................................................',
	'......................#.............................................................',
	'......................#.............................................................',
];

function worldMap(slide, x, y, w, h) {
	const cw = w / WORLD[0].length;
	const ch = h / WORLD.length;
	WORLD.forEach((row, r) => {
		for (let c = 0; c < row.length; c++) {
			if (row[c] !== '#') continue;
			let end = c;
			while (end + 1 < row.length && row[end + 1] === '#') end++;
			rect(slide, x + c * cw, y + r * ch, (end - c + 1) * cw + 0.008, ch + 0.008, {
				fill: { color: mix(GREEN, TEAL, (c + end) / 2 / row.length) },
			});
			c = end;
		}
	});
}

/** Slide 23 — "World Map". */
function slide23(s) {
	ring(s, 10.612, -2.467, 4.581);
	tx(s, 'World Map', { x: 0.571, y: 0.632, w: 5.0, h: 0.695, fontFace: INTER_SB, fontSize: 36, lineSpacing: 45 });
	tx(s, 'Medical is an important factor. Ability to acquire skills some one.', {
		x: 0.559, y: 1.528, w: 5.715, h: 0.761, fontFace: SANS, fontSize: 18, lineSpacing: 24,
	});

	worldMap(s, 3.24, 2.349, 9.562, 4.672);

	// location pins with counters
	[
		[4.809, 3.172, '10+', 4.657, 3.475], [6.011, 5.332, '3.3M', 5.734, 5.67],
		[8.25, 4.685, '10+', 8.079, 5.016], [10.174, 3.289, '11M', 9.969, 3.647],
		[11.25, 5.724, '2K', 11.132, 6.002],
	].forEach(([px, py, label, lx, ly]) => {
		s.addShape('teardrop', { x: px, y: py, w: 0.233, h: 0.233, fill: { color: BLACK }, rotate: 135 });
		ellipse(s, px + 0.063, py + 0.065, 0.105, 0.105, { fill: { color: WHITE } });
		tx(s, label, { x: lx, y: ly, w: 0.822, h: 0.337, fontFace: INTER, fontSize: 14, color: WHITE, align: 'center' });
	});

	[
		{ x: 0.601, value: '73%', w: 0.932 }, { x: 3.5, value: '730K', w: 1.215 },
	].forEach((c) => {
		tx(s, c.value, { x: c.x, y: 5.471, w: c.w, h: 0.55, fontFace: INTER_SB, fontSize: 24, lineSpacing: 35 });
		tx(s, 'Your Title Here', { x: c.x, y: 6.003, w: 2.071, h: 0.286, fontFace: INTER_M, fontSize: 11 });
		tx(s, 'There are many variations of passages', {
			x: c.x, y: 6.375, w: 1.915, h: 0.527, fontFace: INTER, fontSize: 8, lineSpacing: 16,
		});
	});
	s.addShape('line', {
		x: 2.862, y: 6.005, w: 0, h: 0.675, line: { color: 'A6A6A6', width: 1, dashType: 'dashDot' },
	});
}

/* ------------------------------------------------------------------ *
 * Build
 * ------------------------------------------------------------------ */

const BUILDERS = [
	slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10, slide11, slide12,
	slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
];

function build() {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'MEDICAL_16x9', width: 13.333, height: 7.5 });
	pptx.layout = 'MEDICAL_16x9';
	pptx.theme = { headFontFace: SANS_SB, bodyFontFace: SANS };

	BUILDERS.forEach((fn) => {
		const slide = pptx.addSlide();
		slide.background = { color: WHITE };
		fn(slide);
	});

	return pptx.writeFile({
		fileName: path.join(__dirname, '0658b937-7af8-446c-9851-b46f924081e5_grok_final.pptx'),
	});
}

build().then((f) => console.log('wrote', f));
