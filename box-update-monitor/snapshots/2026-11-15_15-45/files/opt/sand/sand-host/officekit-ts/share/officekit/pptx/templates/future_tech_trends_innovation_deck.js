/**
 * "TechNology" deck — rebuilt with pptxgenjs.
 * Slide size 13.333 x 7.5 in (16:9). Dark theme, purple/yellow/green accents.
 * Photographs in the original are replaced by flat grey placeholder shapes.
 */
'use strict';

const PptxGenJS = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ *
 * Palette & typography (from the deck theme)
 * ------------------------------------------------------------------ */
const BLACK = '000000'; // theme bg1
const WHITE = 'FFFFFF'; // theme tx1 / bg2
const PURPLE = 'AD74FE'; // theme accent1
const YELLOW = 'F4C606'; // theme accent2
const GREEN = '26BD98'; // theme accent3
const PURPLE_DK = '7718FD'; // accent1 lumMod 75%
const WAVE = '0D0D0D'; // bg1 lumMod 95% / lumOff 5%
const CARD = '262626'; // bg1 lumMod 85% / lumOff 15%
const PHOTO = 'CFCFCF'; // stand-in for the stock photos
const PHOTO_TXT = 'AFAFAF';
const SILVER = 'C9CACE';
const HEAD = 'Red Hat Display'; // theme major font
const BODY = 'Open Sans'; // theme minor font

/* ------------------------------------------------------------------ *
 * Small drawing helpers
 * ------------------------------------------------------------------ */

/** Heading / display text (Red Hat Display, top anchored like the original). */
function head(slide, text, o) {
	slide.addText(text, Object.assign({ fontFace: HEAD, color: WHITE, valign: 'top' }, o));
}

/** Body copy (Open Sans, 10pt, 150% leading unless overridden). */
function body(slide, text, o) {
	slide.addText(
		text,
		Object.assign({ fontFace: BODY, fontSize: 10, color: WHITE, valign: 'top', lineSpacingMultiple: 1.5 }, o)
	);
}

/** Rounded rectangle. */
function box(slide, o) {
	slide.addShape('roundRect', Object.assign({ line: { width: 0 } }, o));
}

/** 8-pointed decorative star (yellow ones are 0.96", green ones 1.274"). */
function star(slide, x, y, size, color) {
	slide.addShape('star8', {
		x: x,
		y: y,
		w: size,
		h: size,
		rectRadius: 0.22855 * size, // sharp points: prstGeom adj = 22855
		fill: { color: color },
		line: { width: 0 },
	});
}

/** Line-drawn glyph: no fill, thin coloured stroke. */
function glyph(slide, shape, o) {
	slide.addShape(shape, Object.assign({ fill: { type: 'none' }, line: { color: WHITE, width: 1 } }, o));
}

/**
 * Rounded-corner path through a polygon of corner points, relative to the
 * shape origin. Works for concave corners too (used for the notched photos).
 */
function roundPoly(corners, r) {
	const K = 0.5523; // circle-through-bezier constant
	const n = corners.length;
	const towards = (c, p) => {
		const dx = p[0] - c[0];
		const dy = p[1] - c[1];
		const d = Math.hypot(dx, dy) || 1;
		return [c[0] + (dx / d) * r, c[1] + (dy / d) * r];
	};
	const pts = [];
	for (let i = 0; i < n; i++) {
		const c = corners[i];
		const a = towards(c, corners[(i - 1 + n) % n]);
		const b = towards(c, corners[(i + 1) % n]);
		if (i === 0) pts.push({ x: a[0], y: a[1], moveTo: true });
		else pts.push({ x: a[0], y: a[1] });
		pts.push({
			x: b[0],
			y: b[1],
			curve: {
				type: 'cubic',
				x1: a[0] + (c[0] - a[0]) * K,
				y1: a[1] + (c[1] - a[1]) * K,
				x2: b[0] + (c[0] - b[0]) * K,
				y2: b[1] + (c[1] - b[1]) * K,
			},
		});
	}
	pts.push({ close: true });
	return pts;
}

/** Grey stand-in for a photo: plain rectangle + "[image]" caption. */
function photo(slide, x, y, w, h, r) {
	if (r) box(slide, { x: x, y: y, w: w, h: h, rectRadius: r, fill: { color: PHOTO } });
	else slide.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: PHOTO }, line: { width: 0 } });
	caption(slide, x + w / 2, y + h / 2);
}

/**
 * Grey stand-in whose outline is a rounded polygon (the notched photos).
 * `cap` is the caption centre, expressed relative to the shape origin.
 */
function photoPoly(slide, x, y, w, h, corners, r, cap) {
	slide.addShape('custGeom', {
		x: x,
		y: y,
		w: w,
		h: h,
		points: roundPoly(corners, r),
		fill: { color: PHOTO },
		line: { width: 0 },
	});
	caption(slide, x + cap[0], y + cap[1]);
}

function caption(slide, cx, cy) {
	slide.addText('[image]', {
		x: cx - 0.75,
		y: cy - 0.18,
		w: 1.5,
		h: 0.36,
		align: 'center',
		valign: 'middle',
		fontFace: BODY,
		fontSize: 11,
		color: PHOTO_TXT,
	});
}

/* ------------------------------------------------------------------ *
 * Background "ribbon" — one hand-drawn squiggle reused on every slide
 * at different sizes / rotations. Points are the unit-square outline.
 * ------------------------------------------------------------------ */
const RIBBON_START = [0.948, 1.0];
const RIBBON_CURVES = [
	[0.935, 1.0, 0.9221, 0.9965, 0.912, 0.9895],
	[0.8176, 0.9242, 0.8411, 0.821, 0.8566, 0.7527],
	[0.8697, 0.6954, 0.8726, 0.6214, 0.8119, 0.5944],
	[0.7812, 0.5807, 0.7362, 0.5797, 0.6887, 0.5786],
	[0.6199, 0.5771, 0.5419, 0.5754, 0.4869, 0.532],
	[0.418, 0.4779, 0.4429, 0.4097, 0.4649, 0.3495],
	[0.4812, 0.3049, 0.4966, 0.2628, 0.4734, 0.2343],
	[0.4485, 0.2037, 0.3828, 0.1954, 0.3458, 0.1932],
	[0.3281, 0.1922, 0.31, 0.1918, 0.2909, 0.1913],
	[0.2415, 0.1902, 0.1904, 0.189, 0.1394, 0.1765],
	[0.0555, 0.1558, -0.0181, 0.0948, 0.004, 0.0287],
	[0.0106, 0.0086, 0.0387, -0.0038, 0.0666, 0.001],
	[0.0945, 0.0058, 0.1118, 0.026, 0.1051, 0.0461],
	[0.0988, 0.0648, 0.1252, 0.094, 0.1731, 0.1058],
	[0.2093, 0.1147, 0.2505, 0.1157, 0.2942, 0.1167],
	[0.3138, 0.1171, 0.334, 0.1176, 0.3544, 0.1188],
	[0.4544, 0.1247, 0.5266, 0.1517, 0.5632, 0.1966],
	[0.6081, 0.2519, 0.5854, 0.3138, 0.5655, 0.3684],
	[0.5452, 0.4239, 0.535, 0.4591, 0.5636, 0.4817],
	[0.589, 0.5017, 0.6367, 0.5027, 0.6919, 0.5039],
	[0.7483, 0.5052, 0.8123, 0.5066, 0.8668, 0.5309],
	[0.9546, 0.5701, 0.9857, 0.6488, 0.9592, 0.7647],
	[0.9433, 0.8347, 0.9332, 0.9005, 0.9841, 0.9357],
	[1.0048, 0.9501, 1.0054, 0.9737, 0.9854, 0.9885],
	[0.9753, 0.9962, 0.9617, 1.0, 0.948, 1.0],
];

/** Placement of the ribbon per slide: [x, y, w, h, rotation, flipH, flipV]. */
const RIBBONS = {
	1: [8.192, -1.436, 5.817, 8.097, 15, true, false],
	2: [4.97, -2.504, 8.334, 11.601, 30, true, false],
	3: [7.033, -0.158, 6.254, 8.705, 15, true, false],
	4: [-1.029, -1.213, 7.542, 10.498, 345, false, false],
	5: [7.033, -0.158, 6.254, 8.705, 15, true, false],
	6: [7.033, -0.158, 6.254, 8.705, 15, true, false],
	7: [1.47, -2.323, 8.334, 11.601, 330, false, false],
	8: [6.396, 0.496, 6.512, 9.064, 90, true, false],
	9: [4.97, -2.504, 8.334, 11.601, 30, true, false],
	10: [4.97, -2.504, 8.334, 11.601, 30, true, false],
	11: [-1.029, -1.213, 7.542, 10.498, 345, false, false],
	12: [6.396, 0.496, 6.512, 9.064, 90, true, false],
	13: [-1.158, -0.232, 6.75, 9.395, 15, false, true],
	14: [-1.029, -1.213, 7.542, 10.498, 345, false, false],
	15: [8.192, -1.436, 5.817, 8.097, 15, true, false],
};

function ribbon(slide, n) {
	const [x, y, w, h, rot, flipH, flipV] = RIBBONS[n];
	const pts = [{ x: RIBBON_START[0] * w, y: RIBBON_START[1] * h, moveTo: true }];
	RIBBON_CURVES.forEach(c => {
		pts.push({
			x: c[4] * w,
			y: c[5] * h,
			curve: { type: 'cubic', x1: c[0] * w, y1: c[1] * h, x2: c[2] * w, y2: c[3] * h },
		});
	});
	pts.push({ close: true });
	slide.addShape('custGeom', {
		x: x,
		y: y,
		w: w,
		h: h,
		rotate: rot,
		flipH: flipH,
		flipV: flipV,
		points: pts,
		fill: { color: WAVE },
		line: { width: 0 },
	});
}

/* ------------------------------------------------------------------ *
 * Shared page furniture: logo, top nav, footer URL, page number
 * ------------------------------------------------------------------ */
const NAV = [
	{ label: 'Home', x: 8.575, bold: true },
	{ label: 'About Us', x: 10.16, bold: false },
	{ label: 'Experience', x: 11.745, bold: false },
];

function chrome(slide, n) {
	head(slide, [{ text: 'Tech', options: { bold: true } }, { text: 'Nology' }], {
		x: 0.383,
		y: 0.453,
		w: 1.978,
		h: 0.303,
		fontSize: 12,
	});
	NAV.forEach(item => {
		body(slide, item.label, {
			x: item.x,
			y: 0.473,
			w: 1.277,
			h: 0.269,
			align: 'center',
			valign: 'bottom',
			bold: item.bold,
			charSpacing: 1,
			lineSpacingMultiple: 1,
		});
	});
	body(slide, 'www.reallygreatsite.com', {
		x: 0.394,
		y: 6.791,
		w: 2.731,
		h: 0.269,
		valign: 'middle',
		charSpacing: 1,
		lineSpacingMultiple: 1,
	});
	slide.slideNumber = {
		x: n < 10 ? 12.637 : 12.312,
		y: 6.717,
		w: n < 10 ? 0.287 : 0.612,
		h: 0.399,
		align: 'right',
		valign: 'middle',
		bold: true,
		fontFace: BODY,
		fontSize: 12,
		color: WHITE,
	};
}

/** Every slide starts the same way: black page, ribbon, chrome. */
function newSlide(pres, n) {
	const slide = pres.addSlide();
	slide.background = { color: BLACK };
	ribbon(slide, n);
	chrome(slide, n);
	return slide;
}

/* ------------------------------------------------------------------ *
 * Icon stand-ins (the originals are small SVG/PNG glyphs)
 * ------------------------------------------------------------------ */
/**
 * Cloud outline centred on (cx, cy). `mark` picks the symbol inside it:
 * 'search' (slide 3) or 'sync' (the up/down arrow on slide 6). `bg` is the
 * colour behind the icon, so the arrow can knock out the cloud outline.
 */
function iconCloud(slide, cx, cy, w, color, mark, bg) {
	glyph(slide, 'cloud', { x: cx - w / 2, y: cy - w * 0.38, w: w, h: w * 0.76, line: { color: color, width: 0.75 } });
	if (mark === 'sync') {
		slide.addShape('upDownArrow', {
			x: cx - w * 0.13,
			y: cy - w * 0.42,
			w: w * 0.26,
			h: w * 0.84,
			fill: { color: bg },
			line: { color: color, width: 0.75 },
		});
		return;
	}
	glyph(slide, 'ellipse', {
		x: cx - w * 0.03,
		y: cy - w * 0.03,
		w: w * 0.28,
		h: w * 0.28,
		line: { color: color, width: 1 },
	});
	slide.addShape('line', {
		x: cx + w * 0.23,
		y: cy + w * 0.23,
		w: w * 0.17,
		h: w * 0.17,
		line: { color: color, width: 1 },
	});
}

/** Sheet of paper with ruled lines. */
function iconDoc(slide, cx, cy, w, color) {
	const h = w * 1.3;
	glyph(slide, 'snip1Rect', {
		x: cx - w / 2,
		y: cy - h / 2,
		w: w,
		h: h,
		flipH: true,
		rectRadius: w * 0.28,
		line: { color: color, width: 1 },
	});
	[0.2, 0.4, 0.6, 0.78].forEach(t => {
		slide.addShape('line', {
			x: cx - w * 0.28,
			y: cy - h / 2 + h * t,
			w: w * 0.56,
			h: 0,
			line: { color: color, width: 1 },
		});
	});
}

/** Stacked-disc database symbol: three cans of decreasing height. */
function iconDatabase(slide, cx, cy, w, color) {
	const h = w * 1.1;
	const band = h * 0.3;
	[0, 1, 2].forEach(i => {
		glyph(slide, 'can', {
			x: cx - w / 2,
			y: cy - h / 2 + i * band,
			w: w,
			h: h - i * band,
			line: { color: color, width: 1 },
		});
	});
}

function iconPin(slide, cx, cy, s, color) {
	slide.addShape('teardrop', {
		x: cx - s / 2,
		y: cy - s / 2,
		w: s,
		h: s,
		rotate: 135,
		fill: { color: color },
		line: { width: 0 },
	});
	slide.addShape('ellipse', {
		x: cx - s * 0.14,
		y: cy - s * 0.2,
		w: s * 0.28,
		h: s * 0.28,
		fill: { color: PURPLE },
		line: { width: 0 },
	});
}

function iconPhone(slide, cx, cy, s, color) {
	slide.addShape('roundRect', {
		x: cx - s * 0.17,
		y: cy - s * 0.5,
		w: s * 0.34,
		h: s,
		rotate: 315,
		rectRadius: s * 0.15,
		fill: { color: color },
		line: { width: 0 },
	});
}

function iconGlobe(slide, cx, cy, s, color) {
	glyph(slide, 'ellipse', { x: cx - s / 2, y: cy - s / 2, w: s, h: s, line: { color: color, width: 1 } });
	glyph(slide, 'ellipse', {
		x: cx - s * 0.22,
		y: cy - s / 2,
		w: s * 0.44,
		h: s,
		line: { color: color, width: 0.75 },
	});
	slide.addShape('line', { x: cx - s / 2, y: cy, w: s, h: 0, line: { color: color, width: 0.75 } });
}

/** Open envelope: white body with a dark V-flap. */
function iconMail(slide, cx, cy, s, color) {
	slide.addShape('rect', {
		x: cx - s / 2,
		y: cy - s * 0.32,
		w: s,
		h: s * 0.64,
		fill: { color: color },
		line: { width: 0 },
	});
	slide.addShape('triangle', {
		x: cx - s * 0.34,
		y: cy - s * 0.3,
		w: s * 0.68,
		h: s * 0.3,
		rotate: 180,
		fill: { color: CARD },
		line: { width: 0 },
	});
}

/** Paper plane pointing up-right, drawn as a small dart outline. */
function iconSend(slide, cx, cy, s, color) {
	const p = [
		[1.0, 0.0],
		[0.42, 1.0],
		[0.3, 0.58],
		[0.0, 0.42],
	];
	slide.addShape('custGeom', {
		x: cx - s / 2,
		y: cy - s / 2,
		w: s,
		h: s,
		points: p
			.map((q, i) => ({ x: q[0] * s, y: q[1] * s, moveTo: i === 0 }))
			.concat([{ close: true }]),
		fill: { color: color },
		line: { width: 0 },
	});
}

/* ------------------------------------------------------------------ *
 * Repeated copy
 * ------------------------------------------------------------------ */
const LOREM =
	'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
	'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit euismod, ' +
	'metus ac Cras euismod, ';
const LOREM_SHORT = 'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, a maximus leo';

/* ------------------------------------------------------------------ *
 * Slide 1 — title / cover
 * ------------------------------------------------------------------ */
const COVER_STATS = [
	{ value: '$1238M', label: 'Tech Evolution', x: 1.529, lx: 1.529, vw: 1.405, lw: 1.168 },
	{ value: '12K+', label: 'Digital Age', x: 2.935, lx: 2.934, vw: 0.845, lw: 0.912 },
	{ value: '$12M+', label: 'Digital Advantage', x: 4.083, lx: 4.083, vw: 1.405, lw: 1.374 },
];

/** Cover layout shared by slide 1 and slide 15 (only title / button label differ). */
function coverSlide(pres, n, title, buttonLabel) {
	const slide = newSlide(pres, n);

	// Purple stat bar, tucked into the notch of the photo above it.
	box(slide, { x: 1.063, y: 4.984, w: 7.325, h: 1.524, rectRadius: 0.254, fill: { color: PURPLE } });

	head(slide, title, { x: 1.086, y: 0.992, w: 5.822, h: 1.919, fontSize: 54 });
	head(slide, 'Innovation in the Digital Age', { x: 7.243, y: 1.356, w: 3.392, h: 0.37, fontSize: 16 });
	body(slide, LOREM, { x: 7.243, y: 1.819, w: 5.004, h: 0.831 });

	COVER_STATS.forEach(s => {
		head(slide, s.value, { x: s.x, y: 5.406, w: s.vw, h: 0.438, fontSize: 20 });
		body(slide, s.label, { x: s.lx, y: 5.814, w: s.lw, h: 0.326 });
	});

	// White pill button with a black circle + send glyph.
	box(slide, { x: 5.632, y: 5.523, w: 2.227, h: 0.467, rectRadius: 0.2335, fill: { color: WHITE } });
	body(slide, buttonLabel, { x: 5.706, y: 5.605, w: 1.73, h: 0.303, fontSize: 12, color: BLACK, lineSpacingMultiple: 1 });
	slide.addShape('ellipse', { x: 7.428, y: 5.582, w: 0.349, h: 0.349, fill: { color: BLACK }, line: { width: 0 } });
	iconSend(slide, 7.6, 5.755, 0.15, WHITE);

	// Wide photo band with an L-shaped notch for the stat bar.
	photoPoly(
		slide,
		1.086,
		3.244,
		11.162,
		3.264,
		[
			[0, 0],
			[11.162, 0],
			[11.162, 3.264],
			[7.455, 3.264],
			[7.455, 1.6],
			[0, 1.6],
		],
		0.25,
		[9.3, 1.6]
	);

	star(slide, 11.165, 0.685, 0.96, YELLOW);
	star(slide, 0.573, 4.288, 1.274, GREEN);
	return slide;
}

/* ------------------------------------------------------------------ *
 * Slide 2 — table of contents
 * ------------------------------------------------------------------ */
const TOC_ROWS = [
	{ num: '01', y: 2.121, numW: 0.909 },
	{ num: '02', y: 3.297, numW: 1.15 },
	{ num: '03', y: 4.472, numW: 1.008 },
];

function slide02(pres) {
	const slide = newSlide(pres, 2);

	TOC_ROWS.forEach(row => {
		box(slide, { x: 8.067, y: row.y, w: 4.117, h: 0.907, rectRadius: 0.151, fill: { color: CARD } });
		head(slide, 'Your Subtitle Goes Here', { x: 9.237, y: row.y + 0.269, w: 3.147, h: 0.37, fontSize: 16 });
		head(slide, row.num, {
			x: 8.228,
			y: row.y + 0.067,
			w: row.numW,
			h: 0.774,
			fontSize: 40,
			bold: true,
			color: WHITE,
			transparency: 95,
		});
	});

	head(slide, 'Table Of', { x: 1.084, y: 2.703, w: 2.491, h: 0.774, fontSize: 40 });
	head(slide, 'Content', { x: 1.084, y: 3.349, w: 4.773, h: 1.447, fontSize: 80 });

	star(slide, 6.98, 0.812, 0.96, YELLOW);
	star(slide, 3.845, 5.157, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 3 — "Smarter Business with Technology" + three feature pills
 * ------------------------------------------------------------------ */
const FEATURES = [
	{ y: 1.636, title: 'Future-Proofing Your Business', color: PURPLE, icon: 'cloud', glyph: WHITE },
	{ y: 3.358, title: 'Navigating the Digital Shift', color: YELLOW, icon: 'doc', glyph: BLACK },
	{ y: 5.08, title: 'Tech-Driven Business Models', color: GREEN, icon: 'db', glyph: BLACK },
];

function slide03(pres) {
	const slide = newSlide(pres, 3);

	FEATURES.forEach(f => {
		box(slide, { x: 7.529, y: f.y, w: 4.821, h: 1.253, rectRadius: 0.626, fill: { color: CARD } });
		slide.addShape('ellipse', {
			x: 7.769,
			y: f.y + 0.224,
			w: 0.806,
			h: 0.806,
			fill: { color: f.color },
			line: { width: 0 },
		});
		const cx = 7.769 + 0.403;
		const cy = f.y + 0.224 + 0.403;
		if (f.icon === 'cloud') iconCloud(slide, cx, cy, 0.42, f.glyph, 'search');
		if (f.icon === 'doc') iconDoc(slide, cx, cy, 0.32, f.glyph);
		if (f.icon === 'db') iconDatabase(slide, cx, cy, 0.38, f.glyph);
		head(slide, f.title, { x: 8.815, y: f.y + 0.261, w: 2.837, h: 0.286, fontSize: 11 });
		body(slide, LOREM_SHORT, { x: 8.815, y: f.y + 0.503, w: 3.31, h: 0.483, fontSize: 8 });
	});

	head(slide, 'Smarter Business with Technology', { x: 0.931, y: 1.603, w: 4.875, h: 1.447, fontSize: 40 });
	head(slide, 'Technology in the Workplace', { x: 0.984, y: 3.572, w: 2.837, h: 0.337, fontSize: 14 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, a maximus ' +
			'leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit euismod, metus ac Cras euismod, ' +
			'metus ac finibus ac finibus finibus, felis dui suscipit purus, a ac finibus finibus, felis dui suscipit',
		{ x: 0.984, y: 3.984, w: 5.683, h: 1.084 }
	);
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
			'a maximus leo ligula at dolor. Morbi et malesuada purus',
		{ x: 0.984, y: 5.165, w: 5.683, h: 0.579 }
	);

	star(slide, 11.625, 1.245, 0.96, YELLOW);
	star(slide, 3.333, 5.916, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 4 — "The Rise of Quantum Computing"
 * ------------------------------------------------------------------ */
function slide04(pres) {
	const slide = newSlide(pres, 4);

	// Tall photo with a circular bite on its right edge.
	photo(slide, 0.958, 1.375, 4.333, 4.75, 0.929);
	slide.addShape('ellipse', { x: 3.847, y: 3.517, w: 1.471, h: 1.471, fill: { color: BLACK }, line: { width: 0 } });
	slide.addShape('ellipse', { x: 4.164, y: 3.773, w: 0.96, h: 0.96, fill: { color: YELLOW }, line: { width: 0 } });
	iconDatabase(slide, 4.644, 4.253, 0.427, BLACK);

	head(slide, 'The Rise of Quantum Computing', { x: 6.137, y: 1.548, w: 5.918, h: 1.447, fontSize: 40 });
	head(slide, '123,45K+', { x: 6.137, y: 3.291, w: 2.358, h: 0.64, fontSize: 32 });
	body(slide, 'Quisque non elit mauris. Cras', { x: 6.137, y: 3.873, w: 2.29, h: 0.326 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
			'a maximus leo ligula at euismod, metus ac',
		{ x: 6.137, y: 4.363, w: 3.43, h: 0.831 }
	);
	head(slide, '50%', { x: 10.16, y: 3.291, w: 1.213, h: 0.64, fontSize: 32 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
			'a maximus leo ligula at euismod, metus ac',
		{ x: 10.16, y: 3.873, w: 2.29, h: 1.336 }
	);
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, a maxi mus ' +
			'leo ligula at dolor. Morbi et malesuada purus. Phasellus a',
		{ x: 6.137, y: 5.373, w: 6.5, h: 0.579 }
	);

	star(slide, 12.062, 1.724, 0.96, YELLOW);
	star(slide, 4.808, 5.952, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 5 — "Virtual & Augmented Reality in Action"
 * ------------------------------------------------------------------ */
function slide05(pres) {
	const slide = newSlide(pres, 5);

	// Purple panel that overlaps the notch in the photo on the right.
	box(slide, { x: 4.271, y: 4.088, w: 5.689, h: 2.013, rectRadius: 0.336, fill: { color: PURPLE } });

	head(slide, 'Virtual & Augmented Reality in Action', { x: 1.155, y: 1.548, w: 5.918, h: 1.447, fontSize: 40 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
			'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit euismod, ',
		{ x: 1.155, y: 3.171, w: 6.997, h: 0.579 }
	);
	head(slide, '50%', { x: 1.155, y: 4.366, w: 1.213, h: 0.64, fontSize: 32 });
	body(slide, LOREM_SHORT.replace(', a maximus leo', ', a maximus'), { x: 1.155, y: 4.947, w: 2.731, h: 0.831 });
	head(slide, 'Robotics and the Future of Labor', { x: 4.654, y: 4.473, w: 3.318, h: 0.337, fontSize: 14 });
	body(slide, LOREM.replace(/, $/, ''), { x: 4.654, y: 4.886, w: 5.041, h: 0.831 });

	star(slide, 0.195, 1.028, 0.96, YELLOW);
	// Photo notched at bottom-left so the purple panel can sit in front of it.
	photoPoly(
		slide,
		8.575,
		1.503,
		3.318,
		4.599,
		[
			[0, 0],
			[3.318, 0],
			[3.318, 4.599],
			[1.531, 4.599],
			[1.531, 2.479],
			[0, 2.479],
		],
		0.363,
		[1.66, 1.24]
	);
	star(slide, 11.188, 5.301, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 6 — "Disruptive Tech for Modern Technology Enterprises"
 * ------------------------------------------------------------------ */
function slide06(pres) {
	const slide = newSlide(pres, 6);

	head(slide, 'Disruptive Tech for Modern Technology Enterprises', {
		x: 1.155,
		y: 1.394,
		w: 7.095,
		h: 1.447,
		fontSize: 40,
	});
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
			'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit euismod, euismod, metus ac finibus ',
		{ x: 1.155, y: 3.213, w: 7.419, h: 0.579 }
	);

	// Purple call-out button with a white icon tile.
	box(slide, { x: 1.251, y: 4.308, w: 3.505, h: 0.828, rectRadius: 0.138, fill: { color: PURPLE } });
	head(slide, 'The Digital Advantage', { x: 2.121, y: 4.534, w: 2.69, h: 0.37, fontSize: 16 });
	box(slide, { x: 1.45, y: 4.49, w: 0.458, h: 0.458, rectRadius: 0.076, fill: { color: WHITE } });
	iconCloud(slide, 1.679, 4.719, 0.26, BLACK, 'sync', WHITE);

	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
			'a maximus leo ligula at dolor. Morbi et malesuada',
		{ x: 1.155, y: 5.447, w: 5.692, h: 0.579 }
	);
	head(slide, '50%', { x: 9.451, y: 4.816, w: 1.141, h: 0.64, fontSize: 32 });
	body(slide, 'Quisque non elit mauris. Cras euis mod, metus', { x: 10.888, y: 4.827, w: 1.39, h: 0.831 });
	body(slide, 'Quisque non elit', { x: 9.451, y: 5.34, w: 1.39, h: 0.326 });

	star(slide, 0.143, 2.341, 0.96, YELLOW);
	star(slide, 10.471, 5.979, 1.274, GREEN);
	photo(slide, 7.235, 4.181, 1.873, 1.873, 0.312);
	photo(slide, 9.108, 1.206, 2.975, 2.975, 0.496);
}

/* ------------------------------------------------------------------ *
 * Slide 7 — three numbered columns
 * ------------------------------------------------------------------ */
const COLUMNS = [
	{ num: '01', x: 0.954, title: 'Biotech Breakthroughs', numW: 0.909, titleW: 2.51 },
	{ num: '02', x: 5.052, title: 'Tech Ethics and Privacy', numW: 1.409, titleW: 2.51 },
	{ num: '03', x: 9.149, title: 'The Rise of Cybercrime', numW: 1.409, titleW: 2.596 },
];
const COLUMN_COPY =
	'Quisque non elit mauris. Cras euismod, metus ac finibus fini bus, felis dui suscipit purus, a maximus leo ligula euismod, metus';

function slide07(pres) {
	const slide = newSlide(pres, 7);

	COLUMNS.forEach((col, i) => {
		head(slide, col.num, {
			x: col.x,
			y: 4.146,
			w: col.numW,
			h: 0.774,
			fontSize: 40,
			bold: true,
			transparency: 90,
		});
		head(slide, col.title, { x: col.x, y: 4.913, w: col.titleW, h: 0.337, fontSize: 14 });
		body(slide, COLUMN_COPY, { x: col.x, y: 5.201, w: 3.23, h: 0.831 });
		if (i > 0) {
			slide.addShape('line', {
				x: col.x - 0.5,
				y: 4.277,
				w: 0,
				h: 1.77,
				line: { color: WHITE, width: 1, transparency: 70 },
			});
		}
	});

	head(slide, 'Smarter Business with Technology', { x: 0.954, y: 1.454, w: 4.976, h: 1.447, fontSize: 40 });
	head(slide, 'Robotics and the Future of Labor', { x: 6.937, y: 1.586, w: 3.223, h: 0.337, fontSize: 14 });
	body(slide, LOREM + 'metus ac finibus ac finibus', { x: 6.937, y: 1.999, w: 5.577, h: 0.831 });
	body(
		slide,
		LOREM + 'metus ac finibus ac finibus finibus, felis dui suscipit purus, a ac finibus finibus, felis dui suscipit',
		{ x: 0.954, y: 3.231, w: 11.766, h: 0.579 }
	);

	star(slide, 11.419, 1.024, 0.96, YELLOW);
	star(slide, -0.651, 4.444, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 8 — "Hacking, Security & Prevention"
 * ------------------------------------------------------------------ */
const SECURITY_CARDS = [
	{ x: 7.356, w: 2.33, fill: CARD, num: '01', numX: 7.455, numW: 0.888, numColor: BLACK, textX: 7.702, title: 'Tech for Good', color: WHITE },
	{ x: 9.885, w: 2.337, fill: PURPLE, num: '02', numX: 9.984, numW: 1.12, numColor: PURPLE_DK, textX: 10.208, title: 'Security by Design', color: WHITE },
];

function slide08(pres) {
	const slide = newSlide(pres, 8);

	SECURITY_CARDS.forEach(card => {
		box(slide, { x: card.x, y: 1.465, w: card.w, h: 1.521, rectRadius: 0.253, fill: { color: card.fill } });
		head(slide, card.num, {
			x: card.numX,
			y: 1.484,
			w: card.numW,
			h: 0.909,
			fontSize: 48,
			color: card.numColor,
			transparency: 75,
		});
		head(slide, card.title, { x: card.textX, y: 1.772, w: 2.061, h: 0.337, fontSize: 14, color: card.color });
		body(slide, 'Quisque non elit mauri. Cras euismod, met', {
			x: card.textX,
			y: 2.13,
			w: 1.951,
			h: 0.579,
			color: card.color,
		});
	});

	head(slide, 'Hacking, Security & Prevention', { x: 0.954, y: 1.454, w: 4.782, h: 1.447, fontSize: 40 });
	head(slide, 'Navigating Data Regulations', { x: 0.954, y: 3.216, w: 2.96, h: 0.337, fontSize: 14 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, a maximus leo ' +
			'ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna tempor sollicitudin ac finibus finibus, felis dui suscipit ',
		{ x: 0.954, y: 3.599, w: 5.612, h: 0.831 }
	);
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
			'a maximus leo ligula at dolor. Morbi et malesuada purus. ',
		{ x: 0.954, y: 4.522, w: 5.612, h: 0.579 }
	);

	// "Learn More" pill (authored as a tall rounded rect turned on its side).
	box(slide, { x: 1.43, y: 5.029, w: 0.353, h: 1.076, rotate: 90, rectRadius: 0.059, fill: { color: CARD } });
	body(slide, 'Learn More', { x: 1.032, y: 5.381, w: 1.149, h: 0.326, align: 'center' });

	star(slide, 3.384, 5.567, 0.96, YELLOW);
	photo(slide, 7.356, 3.293, 4.866, 2.742, 0.457);
	star(slide, 11.421, 2.896, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 9 — team
 * ------------------------------------------------------------------ */
const TEAM = [
	{ photoX: 1.638, nameX: 0.966, nameW: 3.216, roleX: 1.073, name: 'Benjamin Esteban' },
	{ photoX: 5.73, nameX: 5.141, nameW: 3.052, roleX: 5.166, name: 'Bryan Govanchy' },
	{ photoX: 9.823, nameX: 9.152, nameW: 3.216, roleX: 9.259, name: 'Benjamin Esteban' },
];

function slide09(pres) {
	const slide = newSlide(pres, 9);

	TEAM.forEach(member => {
		photo(slide, member.photoX, 2.814, 1.873, 1.873, 0.312);
		head(slide, member.name, {
			x: member.nameX,
			y: 5.225,
			w: member.nameW,
			h: 0.37,
			fontSize: 16,
			align: 'center',
		});
		body(slide, 'Quisque non elit mauris. Cras', {
			x: member.roleX,
			y: 5.566,
			w: 3.004,
			h: 0.326,
			align: 'center',
		});
	});

	head(slide, 'Meet Our Best Team Here', { x: 1.903, y: 1.454, w: 9.528, h: 0.774, fontSize: 40, align: 'center' });
	star(slide, 8.891, 6.085, 0.96, YELLOW);
	star(slide, 0.317, 1.454, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 10 — break slide
 * ------------------------------------------------------------------ */
function slide10(pres) {
	const slide = newSlide(pres, 10);

	head(slide, 'It\u2019s Time to Break', { x: 1.084, y: 2.353, w: 4.773, h: 2.794, fontSize: 80 });

	box(slide, { x: 7.476, y: 2.167, w: 4.713, h: 3.167, rectRadius: 0.528, fill: { color: PURPLE } });
	head(slide, '30 Minutes Coffee Break', { x: 7.997, y: 2.654, w: 3.671, h: 0.37, fontSize: 16 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
			'a maximus leo ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna tempor sollicitudin',
		{ x: 7.997, y: 3.117, w: 3.748, h: 1.084 }
	);
	box(slide, { x: 8.455, y: 4.096, w: 0.353, h: 1.076, rotate: 90, rectRadius: 0.059, fill: { color: WHITE } });
	body(slide, 'Learn More', { x: 8.056, y: 4.448, w: 1.149, h: 0.326, align: 'center', color: BLACK });

	star(slide, 7.476, 0.933, 0.96, YELLOW);
	star(slide, 1.857, 5.332, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 11 — desktop mock-up
 * ------------------------------------------------------------------ */
function slide11(pres) {
	const slide = newSlide(pres, 11);

	// Desktop computer stand-in: screen, camera dot, chin, neck, foot.
	photo(slide, 1.242, 2.032, 4.844, 2.618, 0);
	slide.addShape('ellipse', { x: 3.64, y: 1.93, w: 0.055, h: 0.055, fill: { color: '4A4A4A' }, line: { width: 0 } });
	box(slide, { x: 1.056, y: 4.806, w: 5.208, h: 0.474, rectRadius: 0.05, fill: { color: SILVER } });
	slide.addShape('trapezoid', {
		x: 3.03,
		y: 5.28,
		w: 1.25,
		h: 0.44,
		flipV: true,
		fill: { color: SILVER },
		line: { width: 0 },
	});
	box(slide, { x: 2.82, y: 5.7, w: 1.67, h: 0.11, rectRadius: 0.05, fill: { color: SILVER } });

	head(slide, 'The Role of Tech in Everyday Life', { x: 7.18, y: 1.888, w: 4.782, h: 1.447, fontSize: 40 });

	// Purple "50%" pill with an arrow.
	box(slide, { x: 7.274, y: 3.75, w: 1.817, h: 0.415, rectRadius: 0.2075, fill: { color: PURPLE } });
	head(slide, '50%', { x: 7.38, y: 3.815, w: 0.668, h: 0.286, fontSize: 10.5 });
	slide.addShape('line', {
		x: 8.182,
		y: 3.953,
		w: 0.693,
		h: 0,
		line: { color: WHITE, width: 1, endArrowType: 'triangle' },
	});

	head(slide, '123,45K+', { x: 7.18, y: 4.324, w: 2.358, h: 0.64, fontSize: 32 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui susci pit purus, a maximus leo ' +
			'ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet metus ac finibus finibus, felis ',
		{ x: 7.18, y: 5.009, w: 5.312, h: 0.831 }
	);

	star(slide, 11.903, 1.531, 0.96, YELLOW);
	star(slide, 4.747, 5.652, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 12 — tablet mock-up
 * ------------------------------------------------------------------ */
function slide12(pres) {
	const slide = newSlide(pres, 12);

	// Tablet stand-in: stylus, rear unit, front unit, screen, home button.
	box(slide, { x: 6.935, y: 1.585, w: 0.155, h: 4.33, rotate: 7, rectRadius: 0.07, fill: { color: 'BEC0BF' } });
	box(slide, { x: 10.29, y: 1.58, w: 1.77, h: 4.36, rectRadius: 0.15, fill: { color: 'E0E2E1' } });
	box(slide, { x: 7.33, y: 1.58, w: 2.96, h: 4.36, rectRadius: 0.18, fill: { color: 'F4F6F7' } });
	photo(slide, 7.481, 2.0, 2.658, 3.51, 0);
	slide.addShape('ellipse', { x: 8.76, y: 1.71, w: 0.05, h: 0.05, fill: { color: '9A9C9B' }, line: { width: 0 } });
	slide.addShape('ellipse', { x: 8.66, y: 5.61, w: 0.16, h: 0.16, fill: { type: 'none' }, line: { color: 'D5D7D6', width: 0.75 } });

	head(slide, 'Redefining Learning with Technology', { x: 0.958, y: 1.888, w: 5.391, h: 1.447, fontSize: 40 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui susci pit purus, a maximus leo ' +
			'ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna tempor sollicitudin. ',
		{ x: 0.958, y: 3.637, w: 5.426, h: 0.831 }
	);
	head(slide, '123,45K+', { x: 0.958, y: 4.755, w: 2.358, h: 0.64, fontSize: 32, fontFace: BODY });
	body(slide, 'Quisque non elit mauris. Cras', { x: 0.958, y: 5.336, w: 2.29, h: 0.326 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, me tus ac finibus finibus, felis dui elit mauris Quisque non elit mauris. ',
		{ x: 3.316, y: 4.831, w: 3.068, h: 0.831 }
	);

	star(slide, 2.81, 0.673, 0.96, YELLOW);
	star(slide, 11.344, 5.317, 1.274, GREEN);
}

/* ------------------------------------------------------------------ *
 * Slide 13 — four labelled notes beside a photo
 * ------------------------------------------------------------------ */
const NOTES = [
	{ title: 'Internet of Things', x: 5.176, y: 4.127, titleW: 1.401, lineX: 6.582, lineY: 4.203, copyX: 6.706, copyY: 4.106 },
	{ title: 'The AI Revolution', x: 5.176, y: 5.267, titleW: 1.677, lineX: 6.582, lineY: 5.343, copyX: 6.706, copyY: 5.245 },
	{ title: 'Automating Success', x: 8.972, y: 4.141, titleW: 1.406, lineX: 10.378, lineY: 4.218, copyX: 10.502, copyY: 4.12 },
	{ title: 'The Digital Advantage', x: 8.972, y: 5.281, titleW: 1.401, lineX: 10.378, lineY: 5.357, copyX: 10.502, copyY: 5.26 },
];

function slide13(pres) {
	const slide = newSlide(pres, 13);

	NOTES.forEach(note => {
		head(slide, note.title, { x: note.x, y: note.y, w: note.titleW, h: 0.572, fontSize: 14 });
		slide.addShape('line', { x: note.lineX, y: note.lineY, w: 0, h: 0.414, line: { color: WHITE, width: 1 } });
		body(slide, 'Quisque non elit mau. Cras euis mod, ', { x: note.copyX, y: note.copyY, w: 1.868, h: 0.579 });
	});

	head(slide, 'Redefining Learning with Technology', { x: 0.958, y: 1.405, w: 5.391, h: 1.447, fontSize: 40 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, a maximus leo ' +
			'ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna tempor sollicitudin ac',
		{ x: 7.354, y: 1.747, w: 5.017, h: 0.831 }
	);
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, a maximus leo ' +
			'ligula at dolor. Morbi et malesuada purus. Phasellus a lacus sit amet urna tempor sollicitudin ac finibus ' +
			'finibus, felis dui suscipit Cras euismod, metus ac finibus finibus, felis ',
		{ x: 0.954, y: 3.074, w: 11.358, h: 0.579 }
	);

	star(slide, 4.691, 0.434, 0.96, YELLOW);
	star(slide, 12.417, 4.063, 1.833, GREEN);
	photo(slide, 0.962, 4.106, 3.593, 1.747, 0.291);
}

/* ------------------------------------------------------------------ *
 * Slide 14 — contact details
 * ------------------------------------------------------------------ */
const CONTACTS = [
	{ x: 6.653, y: 4.412, color: PURPLE, icon: 'pin', textX: 7.213, text: '123 Anywhere St., Any City', glyph: WHITE },
	{ x: 6.659, y: 5.134, color: GREEN, icon: 'phone', textX: 7.213, text: '+123-456-7890', glyph: WHITE },
	{ x: 9.601, y: 4.412, color: YELLOW, icon: 'globe', textX: 10.161, text: 'www.reallygreatsite.com', glyph: WHITE },
	{ x: 9.601, y: 5.134, color: CARD, icon: 'mail', textX: 10.161, text: 'hello@reallygreatsite.com', glyph: WHITE },
];

function slide14(pres) {
	const slide = newSlide(pres, 14);

	photo(slide, 1.004, 1.587, 4.732, 4.325, 0.721);

	head(slide, 'Our Contact Information Here', { x: 6.517, y: 1.928, w: 4.732, h: 1.447, fontSize: 40 });
	body(
		slide,
		'Quisque non elit mauris. Cras euismod, metus ac finibus finibus, felis dui suscipit purus, ' +
			'a maxi mus leo ligula at dolor. Morbi et malesuada',
		{ x: 6.517, y: 3.512, w: 5.839, h: 0.579 }
	);

	CONTACTS.forEach(c => {
		box(slide, { x: c.x, y: c.y, w: 0.438, h: 0.438, rectRadius: 0.073, fill: { color: c.color } });
		const cx = c.x + 0.219;
		const cy = c.y + 0.219;
		if (c.icon === 'pin') iconPin(slide, cx, cy, 0.2, c.glyph);
		if (c.icon === 'phone') iconPhone(slide, cx, cy, 0.24, c.glyph);
		if (c.icon === 'globe') iconGlobe(slide, cx, cy, 0.23, c.glyph);
		if (c.icon === 'mail') iconMail(slide, cx, cy, 0.24, c.glyph);
		body(slide, c.text, { x: c.textX, y: c.y + 0.042, w: 1.977, h: 0.353, valign: 'middle' });
	});

	star(slide, 11.551, 1.508, 0.96, YELLOW);
	star(slide, 0.297, 4.613, 1.833, GREEN);
}

/* ------------------------------------------------------------------ *
 * Build the deck
 * ------------------------------------------------------------------ */
function build() {
	const pres = new PptxGenJS();
	pres.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
	pres.layout = 'DECK';
	pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
	pres.title = 'The Future of Technology';

	coverSlide(pres, 1, 'The Future of Technology', 'Start Presentation');
	slide02(pres);
	slide03(pres);
	slide04(pres);
	slide05(pres);
	slide06(pres);
	slide07(pres);
	slide08(pres);
	slide09(pres);
	slide10(pres);
	slide11(pres);
	slide12(pres);
	slide13(pres);
	slide14(pres);
	coverSlide(pres, 15, 'Thanks for Your Attention to Us', 'End Presentation');

	return pres.writeFile({
		fileName: path.join(__dirname, '0be8e22a-3046-4d34-9069-e108feabc0d0_grok_final.pptx'),
	});
}

build().then(f => console.log('wrote', f));
