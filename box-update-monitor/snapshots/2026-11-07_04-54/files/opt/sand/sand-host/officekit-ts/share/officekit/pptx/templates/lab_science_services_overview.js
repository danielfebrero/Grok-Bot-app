/**
 * Biolatory - Laboratory & Science Presentation (24 slides, 13.333" x 7.5")
 * Rebuilt with pptxgenjs. Raster artwork from the source deck is replaced by
 * native-shape placeholders (see `iconMark` / `imagePlaceholder`).
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const RED = 'E31B29';
const WHITE = 'FFFFFF';
const BLACK = '000000';
const GRAY = '595959'; // tx1 lum 65/35 - body copy
const GRAY_DARK = '262626'; // tx1 lum 85/15
const GRAY_XDARK = '0D0D0D'; // tx1 lum 95/5
const GRAY_MID = '7F7F7F'; // tx1 lum 50/50
const GRAY_LT = 'F2F2F2'; // bg1 lum 95
const GRAY_LINE = 'D9D9D9'; // bg1 lum 85
const CREAM = 'F9F8F3';
const AMBER = 'FFC000';

/* -------------------------------------------------------------------- fonts */
const JK = 'Plus Jakarta Sans';
const JK_SB = 'Plus Jakarta Sans SemiBold';
const OPEN = 'Open Sans';
const POPPINS = 'Poppins';
const POPPINS_SB = 'Poppins SemiBold';
const MULISH = 'Mulish';
const INTER_MED = 'Inter Medium';

/* ------------------------------------------------------------------ shadows */
// Offsets are in points. The "glow" presets are centred (offset 0 in the source
// deck) but pptxgenjs treats a literal 0 as "unset", so they use a hair above 0.
// NOTE: pptxgenjs rewrites the shadow object it is handed, so every shape must
// get its own copy - hence `shadow()` returns a new object each call.
const SHADOWS = {
	glow: [35, 0.01, 90, BLACK, 0.15],
	glowSoft: [35, 0.01, 90, BLACK, 0.1],
	glowFaint: [35, 0.01, 90, BLACK, 0.06],
	lift: [25, 3, 90, BLACK, 0.1],
	liftStrong: [25, 3, 90, BLACK, 0.2],
	drop: [55, 15, 45, BLACK, 0.15],
	dropCool: [55, 15, 45, '2C3138', 0.11],
	dropTiny: [12, 9, 45, BLACK, 0.14],
	dropWide: [18, 9, 45, BLACK, 0.05],
	halo: [10, 4, 90, BLACK, 0.15],
	haloSoft: [10, 4, 90, BLACK, 0.11],
	haloWide: [17, 0.01, 90, BLACK, 0.05],
	panel: [30, 5, 90, BLACK, 0.13],
};
function shadow(name) {
	const [blur, offset, angle, color, opacity] = SHADOWS[name];
	return { type: 'outer', blur, offset, angle, color, opacity };
}

/* ------------------------------------------------------------ shape helpers */
let S = null; // pptxgenjs ShapeType map, assigned in build()

/** filled rectangle */
function rect(slide, x, y, w, h, opts = {}) {
	slide.addShape(S.rect, Object.assign({ x, y, w, h }, opts));
}

/** filled ellipse */
function oval(slide, x, y, w, h, opts = {}) {
	slide.addShape(S.ellipse, Object.assign({ x, y, w, h }, opts));
}

/** fully rounded "pill" (corner radius = half the height) */
function pill(slide, x, y, w, h, opts = {}) {
	slide.addShape(S.roundRect, Object.assign({ x, y, w, h, rectRadius: h / 2 }, opts));
}

/** horizontal rule / arrow */
function hline(slide, x, y, w, opts = {}) {
	slide.addShape(S.line, Object.assign({ x, y, w, h: 0 }, opts));
}

/** text block; anchors top-left like the source deck's text boxes */
function txt(slide, text, o) {
	slide.addText(text, Object.assign({ fontFace: JK, fontSize: 18, color: BLACK, valign: 'top' }, o));
}

/** 11pt Open Sans body copy at 1.5 line spacing */
function body(slide, text, o) {
	txt(slide, text, Object.assign({ fontFace: OPEN, fontSize: 11, color: GRAY, lineSpacingMultiple: 1.5 }, o));
}

/** button: colored rectangle with centered label */
function button(slide, x, y, w, h, label, o = {}) {
	slide.addText(label, Object.assign({
		shape: S.rect, x, y, w, h,
		fill: { color: o.fill || RED }, color: o.color || WHITE,
		fontFace: o.fontFace || JK, fontSize: o.fontSize || 14,
		align: 'center', valign: 'middle',
	}, o.shadow ? { shadow: o.shadow } : {}));
}

/** placeholder standing in for a small raster icon: ring + centre dot */
function iconMark(slide, x, y, w, h, color) {
	oval(slide, x, y, w, h, { fill: { type: 'none' }, line: { color, width: 1.5 } });
	oval(slide, x + w * 0.38, y + h * 0.38, w * 0.24, h * 0.24, { fill: { color } });
}

/** placeholder standing in for a photograph / device mock-up */
function imagePlaceholder(slide, x, y, w, h, o = {}) {
	slide.addText('[image]', Object.assign({
		shape: o.rounded ? S.roundRect : S.rect, x, y, w, h,
		fill: { color: o.fill || GRAY_LT }, line: { color: o.line || GRAY_LINE, width: o.lineWidth || 1.25 },
		color: GRAY_MID, fontFace: OPEN, fontSize: 11, align: 'center', valign: 'middle',
	}, o.rounded ? { rectRadius: o.rounded } : {}, o.rotate ? { rotate: o.rotate } : {}));
}

/**
 * The Biolatory mark: an outlined conical flask tipped 45 degrees clockwise,
 * with bubbles escaping the mouth. The outline is written upright (unit box,
 * mouth at the top) and every vertex is then rotated about the box centre.
 * Bubble offsets are fractions of the mark's bounding box.
 */
const FLASK_OUTLINE = [
	[0.36, 0.0], [0.64, 0.0], // mouth
	[0.64, 0.30], [1.0, 0.90], // right shoulder down to the base
	[0.92, 1.0], [0.08, 1.0], [0.0, 0.90], // base
	[0.36, 0.30], // left shoulder back up to the neck
];
const FLASK_BUBBLES = [
	{ x: 0.878, y: 0.023, d: 0.116, solid: false }, // escaping the mouth
	{ x: 0.888, y: 0.219, d: 0.039, solid: true },
	{ x: 0.400, y: 0.545, d: 0.116, solid: false }, // suspended inside
	{ x: 0.235, y: 0.485, d: 0.083, solid: false },
	{ x: 0.330, y: 0.700, d: 0.044, solid: true },
];
function flaskLogo(slide, x, y, w, h, color) {
	const stroke = { fill: { type: 'none' }, line: { color, width: 4 } };
	const a = 45 * Math.PI / 180, cos = Math.cos(a), sin = Math.sin(a);
	const fw = w * 0.86, fh = h * 0.94, scale = 0.78;
	const spin = ([px, py]) => ({
		x: fw * (0.5 + ((px - 0.5) * cos - (py - 0.5) * sin) * scale),
		y: fh * (0.5 + ((px - 0.5) * sin + (py - 0.5) * cos) * scale),
	});
	slide.addShape(S.custGeom, Object.assign({
		x, y: y + h * 0.06, w: fw, h: fh,
		points: FLASK_OUTLINE.map(spin).concat([{ close: true }]),
	}, stroke));
	FLASK_BUBBLES.forEach((b) => {
		oval(slide, x + w * b.x, y + h * b.y, w * b.d, h * b.d, b.solid ? { fill: { color } } : stroke);
	});
}

/** check-mark glyph drawn as an open polyline */
function checkMark(slide, x, y, w, h, color) {
	slide.addShape(S.custGeom, {
		x, y, w, h,
		points: [{ x: 0, y: h * 0.45 }, { x: w * 0.38, y: h }, { x: w, y: 0 }],
		fill: { type: 'none' }, line: { color, width: 1.5 },
	});
}

/** diagonal red wedge used by the "device mock-up" layouts (slides 21 & 22) */
function redWedge(slide, x, mirrored) {
	const w = 5.648, h = 7.5, top = 3.032;
	const pts = mirrored
		? [{ x: w, y: 0 }, { x: w - top, y: 0 }, { x: 0, y: h }, { x: w, y: h }]
		: [{ x: 0, y: 0 }, { x: top, y: 0 }, { x: w, y: h }, { x: 0, y: h }];
	slide.addShape(S.custGeom, { x, y: 0, w, h, points: pts.concat([{ close: true }]), fill: { color: RED } });
}

/** banner whose two top corners are full half-height semicircles (slide 9) */
function topRoundedBanner(slide, x, y, w, h, opts) {
	const r = h / 2;
	slide.addShape(S.custGeom, Object.assign({
		x, y, w, h,
		points: [
			{ x: 0, y: h }, { x: 0, y: r },
			{ curve: { type: 'arc', hR: r, wR: r, stAng: 180, swAng: 90 }, x: r, y: 0 },
			{ x: w - r, y: 0 },
			{ curve: { type: 'arc', hR: r, wR: r, stAng: 270, swAng: 90 }, x: w, y: r },
			{ x: w, y: h }, { close: true },
		],
	}, opts));
}

/** donut gauge (native chart) - source deck uses an 85% hole, 90/40 split */
function donut(slide, pres, x, y, size, arcColor) {
	slide.addChart(pres.ChartType.doughnut,
		[{ name: 'Region 1', labels: ['April', 'May'], values: [90, 40] }],
		{
			x, y, w: size, h: size,
			holeSize: 85, chartColors: [arcColor, GRAY_LT],
			showLegend: false, showValue: false, showTitle: false,
			dataBorder: { pct: 0, color: WHITE },
			chartArea: { fill: { color: WHITE, transparency: 100 } },
		});
}

/* ============================================================ slide builders */

// 1 - Title: full-bleed red cover
function slide01(pres) {
	const s = pres.addSlide();
	rect(s, 0, 0, 13.333, 7.5, { fill: { color: RED, transparency: 15 } });
	txt(s, 'BIOLATORY', {
		x: 1.745, y: 2.557, w: 9.844, h: 1.717,
		fontSize: 96, bold: true, color: WHITE, align: 'center', charSpacing: 3,
	});
	txt(s, 'LABORATORY & SCIENCE PRESENTATION', {
		x: 3.317, y: 4.274, w: 6.7, h: 0.37,
		fontSize: 16, color: WHITE, align: 'center', charSpacing: 1,
	});
	// top navigation
	txt(s, 'ABOUT US', {
		shape: S.rect, x: 11.587, y: 0.448, w: 1.384, h: 0.508,
		fill: { color: WHITE }, shadow: shadow('lift'),
		fontSize: 12, align: 'center', valign: 'middle',
	});
	txt(s, 'SERVICES', { x: 9.972, y: 0.551, w: 1.514, h: 0.303, fontSize: 12, color: WHITE, align: 'center' });
	txt(s, 'CONTACT US', { x: 8.458, y: 0.551, w: 1.514, h: 0.303, fontSize: 12, color: WHITE, align: 'center' });
	hline(s, 2.983, 6.89, 10.35, { line: { color: WHITE, width: 1 } });
	txt(s, 'Excellence in Laboratory Services', { x: 0.362, y: 6.604, w: 2.503, h: 0.572, fontSize: 14, color: WHITE });
}

// 2 - Intro: logo tile + headline
function slide02(pres) {
	const s = pres.addSlide();
	rect(s, 3.383, 0, 3.221, 2.6, { fill: { color: RED, transparency: 15 } });
	flaskLogo(s, 4.391, 0.388, 1.205, 1.2, WHITE);
	txt(s, 'BIOLATORY', {
		x: 3.741, y: 1.774, w: 2.506, h: 0.438,
		fontSize: 20, bold: true, color: WHITE, align: 'center', charSpacing: 2,
	});
	txt(s, [
		{ text: 'Biolatory The ', options: { breakLine: true } },
		{ text: 'Future of Science And Technology.' },
	], { x: 7.393, y: 0.993, w: 5.229, h: 2.121, fontSize: 40, bold: true });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed tortor eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim mollit minim veniam, quis nostrud exercitation ullamco laboris nisi sed aliquip deserunt commodo consequat irure dolor.',
		{ x: 7.393, y: 3.339, w: 5.229, h: 1.182 });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed tortor eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim mollit minim veniam, quis nostrud exercitation',
		{ x: 7.393, y: 4.718, w: 5.229, h: 0.904 });
	button(s, 7.486, 5.937, 1.791, 0.569, 'Learn More');
}

// 3 - Statement over a red footer band
function slide03(pres) {
	const s = pres.addSlide();
	rect(s, 0, 5.104, 13.333, 2.396, { fill: { color: RED } }); // from layout
	const blurb = 'Lorem ipsum dolor sit amet consectetur adipiscing tortor eiusmod tempor incididunt enim labore non dolore magna aliqua. ';
	[2.616, 8.603].forEach((x) => {
		body(s, blurb, { x, y: 6.0, w: 3.949, h: 0.904, color: WHITE });
		txt(s, 'Simple Text Here', { x, y: 5.697, w: 2.625, h: 0.337, fontFace: JK_SB, fontSize: 14, color: WHITE });
	});
	txt(s, 'Accuracy and Efficiency in Laboratory Analysis.', { x: 0.795, y: 0.931, w: 6.076, h: 1.919, fontSize: 36, bold: true });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed tortor eiusmod tempor incididunt enim labore et dolore magna aliqua. Ut enim mollit minim veniam, quis nostrud exercitation ullamco laboris nisi sed aliquip deserunt commodo eiusmod consequat minim veniam irure dolor.',
		{ x: 0.795, y: 2.991, w: 6.076, h: 1.182 });
}

// 4 - Red promo card + three icon/text rows
function slide04(pres) {
	const s = pres.addSlide();
	const rows = [
		{ y: 1.138, title: 'Biology Laboratory', bold: true, iconSize: 0.558 },
		{ y: 3.234, title: 'Medical Laboratory ', bold: false, iconSize: 0.558 },
		{ y: 5.33, title: 'Scientist Laboratory ', bold: true, iconSize: 0.411 },
	];
	rows.forEach((r) => {
		oval(s, 8.405, r.y, 1.032, 1.032, { fill: { color: WHITE }, shadow: shadow('glow') });
		oval(s, 8.486, r.y + 0.081, 0.87, 0.87, { fill: { color: RED }, shadow: shadow('glowFaint') });
		const pad = (0.87 - r.iconSize) / 2;
		iconMark(s, 8.486 + pad, r.y + 0.081 + pad, r.iconSize, r.iconSize, WHITE);
	});
	rows.forEach((r) => {
		txt(s, r.title, { x: 9.688, y: r.y - 0.116, w: 2.625, h: 0.337, fontFace: JK_SB, fontSize: 14, bold: r.bold });
		body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing enim sed non eiusmod tortor tempor incididunt.',
			{ x: 9.688, y: r.y + 0.221, w: 3.057, h: 0.908 });
	});
	rect(s, 1.029, 1.311, 5.741, 4.878, { fill: { color: RED, transparency: 10 } });
	txt(s, 'Experience The Best Quality Laboratory Products Here.',
		{ x: 1.52, y: 1.832, w: 4.757, h: 1.717, fontSize: 32, bold: true, color: WHITE });
	button(s, 1.614, 5.099, 1.791, 0.569, 'Learn More', { fill: WHITE, color: BLACK, fontFace: JK_SB });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed non eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis tortor nostrud exercitation ullamco laboris nisi ut aliquip enim.',
		{ x: 1.521, y: 3.66, w: 4.757, h: 1.182, color: WHITE });
}

// 5 - Headline + three KPI tiles
function slide05(pres) {
	const s = pres.addSlide();
	rect(s, 7.335, 1.185, 4.22, 5.129, { fill: { color: WHITE }, shadow: shadow('panel') });
	txt(s, 'Leading The Way in Laboratory Innovation and Excellence.',
		{ x: 0.907, y: 1.23, w: 5.521, h: 1.919, fontSize: 35, bold: true });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed tortor eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim mollit minim veniam, quis nostrud exercitation ullamco laboris nisi sed aliquip deserunt commodo consequat duis aute irure dolor.',
		{ x: 0.907, y: 3.285, w: 5.521, h: 1.182 });
	const tiles = [
		{ x: 0.996, fill: RED, value: '15+', valueColor: WHITE, label: 'Experience', labelColor: WHITE },
		{ x: 2.793, fill: WHITE, value: '30', valueColor: RED, label: 'Certification', labelColor: BLACK },
		{ x: 4.59, fill: WHITE, value: '120', valueColor: RED, label: 'Global Partner', labelColor: BLACK },
	];
	tiles.forEach((t) => {
		rect(s, t.x, 4.772, 1.682, 1.497, { fill: { color: t.fill }, shadow: shadow('glow') });
		txt(s, t.value, { x: t.x + 0.379, y: 5.073, w: 0.924, h: 0.606, fontSize: 30, bold: true, color: t.valueColor, align: 'center' });
		txt(s, t.label, { x: t.x + 0.128, y: 5.666, w: 1.425, h: 0.303, fontFace: JK_SB, fontSize: 12, color: t.labelColor, align: 'center' });
	});
}

// 6 - Commitment panel with tick list
function slide06(pres) {
	const s = pres.addSlide();
	txt(s, 'We\u2019re The Lab With The Best Science Experiments.',
		{ x: 0.705, y: 0.699, w: 6.236, h: 1.178, fontSize: 32, bold: true });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit tortor tortor eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim mollit minim veniam quis tortor nostrud exercitation ullamco laboris nisi sed aliquip.',
		{ x: 0.705, y: 1.98, w: 6.236, h: 0.904 });
	rect(s, 7.646, 0, 5.688, 3.583, { fill: { color: RED } });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed non eiusmod tempor incididunt non labore et dolore magna aliqua enim ullamco minim veniam.',
		{ x: 8.229, y: 0.826, w: 4.521, h: 0.904, color: WHITE });
	txt(s, 'We commit to what we provide.',
		{ x: 8.229, y: 0.49, w: 3.878, h: 0.337, fontFace: JK_SB, fontSize: 14, bold: true, color: WHITE });
	['Research Center', 'Cutting-edge Technology', 'Quality Assurance'].forEach((label, i) => {
		const y = 1.899 + i * 0.446;
		txt(s, label, { x: 8.672, y, w: 3.042, h: 0.303, fontFace: JK_SB, fontSize: 12, color: WHITE });
		oval(s, 8.342, y + 0.038, 0.227, 0.227, { fill: { color: WHITE } });
		checkMark(s, 8.413, y + 0.122, 0.085, 0.058, RED);
	});
	const tail = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed non eiusmod tempor incididunt non labore et dolore magna aliqua enim ullamco minim veniam.';
	body(s, tail, { x: 8.229, y: 4.694, w: 4.521, h: 0.904 });
	txt(s, 'Simple Text Here', { x: 8.229, y: 4.357, w: 3.878, h: 0.337, fontFace: JK_SB, fontSize: 14 });
	body(s, tail, { x: 8.229, y: 5.803, w: 4.521, h: 0.904 });
}

// 7 - Red cover with three service cards
function slide07(pres) {
	const s = pres.addSlide();
	rect(s, 0, 0, 13.333, 7.5, { fill: { color: RED, transparency: 20 } });
	[
		{ x: 0.832, title: 'Analytical Laboratory Services' },
		{ x: 4.866, title: 'Materials and Technology Expertise' },
		{ x: 8.903, title: 'Outstanding Laboratory Testing' },
	].forEach((c) => {
		rect(s, c.x, 3.238, 3.601, 3.512, { fill: { color: WHITE }, shadow: shadow('glow') });
		txt(s, c.title, { x: c.x + 0.25, y: 4.732, w: 3.101, h: 0.572, fontFace: POPPINS_SB, fontSize: 14, align: 'center' });
		body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed non eiusmod tempor incididunt labore magna. ',
			{ x: c.x + 0.25, y: 5.394, w: 3.101, h: 0.904, align: 'center' });
		oval(s, c.x + 1.385, 3.69, 0.83, 0.83, { fill: { color: RED } });
		iconMark(s, c.x + 1.534, 3.839, 0.533, 0.533, WHITE);
	});
	txt(s, 'Get The Best Lab Products With Us.',
		{ x: 3.201, y: 0.631, w: 6.931, h: 1.313, fontSize: 36, bold: true, color: WHITE, align: 'center' });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit tortor tortor eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim mollit minim veniam quis tortor.',
		{ x: 3.123, y: 2.013, w: 7.087, h: 0.627, color: WHITE, align: 'center' });
}

// 8 - Commitment headline with two feature cards
function slide08(pres) {
	const s = pres.addSlide();
	const blurb = 'PLACEHOLDER';
	// light card
	rect(s, 3.5, 4.387, 4.417, 2.252, { fill: { color: WHITE }, shadow: shadow('glow') });
	body(s, blurb, { x: 3.801, y: 5.43, w: 3.78, h: 0.904 });
	txt(s, 'Simple Text Here', { x: 4.554, y: 4.847, w: 2.293, h: 0.37, fontFace: JK_SB, fontSize: 16 });
	oval(s, 3.801, 4.689, 0.642, 0.649, { fill: { color: RED }, shadow: shadow('glowFaint') });
	iconMark(s, 3.948, 4.839, 0.349, 0.349, WHITE);
	// red card
	rect(s, 8.218, 4.387, 4.417, 2.252, { fill: { color: RED }, shadow: shadow('glow') });
	oval(s, 8.519, 4.689, 0.642, 0.649, { fill: { color: WHITE }, shadow: shadow('glowFaint') });
	iconMark(s, 8.629, 4.803, 0.421, 0.421, RED);
	body(s, blurb, { x: 8.519, y: 5.43, w: 3.78, h: 0.904, color: GRAY_LT });
	txt(s, 'Simple Text Here', { x: 9.272, y: 4.847, w: 2.293, h: 0.37, fontFace: JK_SB, fontSize: 16, color: WHITE });

	txt(s, 'Commitment to Precise and Timely Lab Testing',
		{ x: 6.538, y: 0.86, w: 6.097, h: 1.919, fontFace: POPPINS_SB, fontSize: 36, bold: true, color: GRAY_XDARK });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed tortor eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim mollit minim veniam sit quis nostrud exercitation ullamco laboris nisi sed aliquip',
		{ x: 6.583, y: 2.89, w: 6.051, h: 0.904 });
}

// 9 - Capability pills on a rounded red banner
function slide09(pres) {
	const s = pres.addSlide();
	topRoundedBanner(s, 0, 3.604, 13.333, 3.896, { fill: { color: RED, transparency: 15 } });
	txt(s, 'The Best Capabilities of Our Laboratory.',
		{ x: 2.984, y: 0.749, w: 7.365, h: 1.313, fontSize: 36, bold: true, align: 'center' });
	[
		{ x: 0.561, y: 5.855, label: 'Sizing Testing' },
		{ x: 2.412, y: 4.437, label: ' Ultimate Analysis' },
		{ x: 7.047, y: 4.437, label: ' Calorific Value' },
		{ x: 4.697, y: 5.855, label: ' Proximate Analysis' },
		{ x: 8.833, y: 5.855, label: 'Medical Testing' },
	].forEach((p) => {
		pill(s, p.x, p.y, 3.94, 0.864, { fill: { color: WHITE }, shadow: shadow('glow') });
		txt(s, p.label, { x: p.x + 0.279, y: p.y + 0.248, w: 2.898, h: 0.37, fontSize: 16, bold: true });
		oval(s, p.x + 3.14, p.y + 0.08, 0.706, 0.706, { fill: { color: RED } });
		hline(s, p.x + 3.337, p.y + 0.433, 0.313, { line: { color: WHITE, width: 1.5, endArrowType: 'arrow' } });
	});
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit tortor tortor eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim mollit minim veniam quis tortor.',
		{ x: 2.002, y: 2.28, w: 9.329, h: 0.627, align: 'center' });
}

// 10 - Meet the team: 2x2 portrait grid
function slide10(pres) {
	const s = pres.addSlide();
	const cells = [
		{ x: 6.326, y: 0.54, fill: WHITE, name: 'Anatasya Theresia ', nameColor: BLACK, roleColor: GRAY },
		{ x: 6.326, y: 3.881, fill: RED, name: 'Nikola Alexander', nameColor: CREAM, roleColor: GRAY_LT },
		{ x: 9.699, y: 3.881, fill: WHITE, name: 'Renata Monalisa', nameColor: BLACK, roleColor: GRAY_MID },
		{ x: 9.699, y: 0.54, fill: RED, name: 'David Bellingham ', nameColor: CREAM, roleColor: GRAY_LT },
	];
	cells.forEach((c) => rect(s, c.x, c.y, 3.059, 3.079, { fill: { color: c.fill }, shadow: shadow('drop') }));
	cells.forEach((c) => {
		txt(s, c.name, { x: c.x - 0.168, y: c.y + 2.167, w: 3.395, h: 0.37, fontSize: 16, bold: true, color: c.nameColor, align: 'center' });
		txt(s, 'Your Position', { x: c.x - 0.168, y: c.y + 2.539, w: 3.395, h: 0.303, fontSize: 12, color: c.roleColor, align: 'center' });
	});
	txt(s, 'Meet Our Team', { x: 0.758, y: 2.276, w: 4.81, h: 0.774, fontSize: 40, bold: true });
	button(s, 0.859, 4.723, 1.757, 0.501, 'Meet Now');
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit sed enim incididunt eiusmod enim tempor incididunt non labore et dolore magna aliqua. Ut enim ad minim veniam tortor nostrud ullamco exercitation laboris nisi ut aliquip.',
		{ x: 0.758, y: 3.197, w: 4.81, h: 1.182 });
}

// 11 - Profile card + skill bars
function slide11(pres) {
	const s = pres.addSlide();
	rect(s, 0, 0, 3.365, 7.5, { fill: { color: RED } }); // from layout
	rect(s, 1.522, 1.099, 3.958, 5.636, { fill: { color: WHITE }, shadow: shadow('dropWide') });
	oval(s, 2.164, 1.367, 2.673, 2.701, { fill: { color: WHITE }, shadow: shadow('halo') });
	txt(s, 'Vanessa Alexandra', { x: 1.786, y: 4.411, w: 3.429, h: 0.438, fontSize: 20, bold: true, align: 'center' });
	txt(s, 'Medical Laboratory', { x: 2.62, y: 4.849, w: 1.762, h: 0.286, fontFace: POPPINS, fontSize: 11, color: GRAY, align: 'center' });
	txt(s, 'Lorem dolor sed viverra ipsum nunc aliquet Bibendum.',
		{ x: 2.051, y: 5.621, w: 2.901, h: 0.627, fontFace: MULISH, fontSize: 11, color: GRAY, align: 'center', lineSpacingMultiple: 1.5 });
	txt(s, 'Detailed Information', { x: 6.996, y: 1.946, w: 3.499, h: 0.368, fontSize: 16, bold: true, color: GRAY_DARK });
	txt(s, 'Lorem dolor sed viverra ipsum nunc aliquet bibendum enim facilisis vestibulum lectus mauris ultrices eros amet noin cursus. Vel quam elementum pulvinar etiam non. Tortor id aliquet lectus proin nibh nisl. Euismod lacinia at quis risus sed vulputate odio. ',
		{ x: 7.0, y: 2.319, w: 4.978, h: 1.179, fontFace: MULISH, fontSize: 11, color: GRAY, align: 'justify', lineSpacingMultiple: 1.5 });
	[
		{ y: 3.751, trackX: 7.249, filled: 4.235, label: '95%' },
		{ y: 4.426, trackX: 7.245, filled: 3.955, label: '90%' },
		{ y: 5.109, trackX: 7.245, filled: 3.672, label: '80%' },
	].forEach((bar) => {
		const barY = bar.y + 0.393;
		pill(s, bar.trackX, barY, 4.664, 0.05, { fill: { color: GRAY_LT } });
		pill(s, 7.087, barY, bar.filled, 0.05, { fill: { color: AMBER } });
		txt(s, 'Your Title Here', { x: 6.996, y: bar.y, w: 3.499, h: 0.286, fontSize: 11, bold: true, color: GRAY_DARK });
		txt(s, bar.label, { x: 11.206, y: bar.y, w: 0.703, h: 0.286, fontSize: 11, bold: true, color: GRAY_DARK, align: 'right' });
	});
}

// 12 - Numbered service pills
function slide12(pres) {
	const s = pres.addSlide();
	rect(s, 7.708, 0, 5.625, 7.5, { fill: { color: GRAY_LT, transparency: 20 } });
	[
		{ y: 1.269, num: '01 |', label: 'Microorganism Chemical', red: true },
		{ y: 2.143, num: '02 |', label: 'Industrial Hygiene', red: false },
		{ y: 3.023, num: '03 |', label: 'Environment', red: true, cornerShadow: true },
		{ y: 3.903, num: '04 |', label: 'Calibration', red: false },
		{ y: 4.778, num: '05 |', label: 'Fertilizers & Pesticides', red: true },
		{ y: 5.652, num: '06 |', label: 'Fundamental Chemistry', red: false },
	].forEach((row) => {
		pill(s, 8.408, row.y, 4.213, 0.579, {
			fill: { color: row.red ? RED : WHITE },
			shadow: shadow(row.cornerShadow ? 'dropTiny' : 'lift'),
		});
		txt(s, row.num, { x: 8.581, y: row.y + 0.087, w: 0.74, h: 0.404, color: row.red ? WHITE : RED, align: 'center' });
		txt(s, row.label, {
			x: 9.101, y: row.y + 0.115, w: 3.346, h: 0.37,
			fontSize: 16, bold: true, color: row.red ? WHITE : BLACK, align: 'center',
		});
	});
	txt(s, 'Science Laboratory Services', { x: 0.816, y: 1.651, w: 6.076, h: 1.447, fontSize: 40, bold: true });
	const para = 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed tortor eiusmod tempor incididunt enim labore et dolore magna aliqua. Ut enim mollit minim veniam, quis nostrud exercitation ullamco laboris nisi sed aliquip deserunt commodo eiusmod consequat minim veniam irure dolor.';
	body(s, para, { x: 0.816, y: 3.225, w: 6.076, h: 1.182 });
	body(s, para, { x: 0.816, y: 4.667, w: 6.076, h: 1.182 });
}

// 13 - Checkerboard of services
function slide13(pres) {
	const s = pres.addSlide();
	const blurb = 'PLACEHOLDER';
	[
		{ x: 1.286, y: 2.389, h: 2.224 },
		{ x: 4.873, y: 4.613, h: 2.246 },
		{ x: 8.46, y: 2.389, h: 2.224 },
	].forEach((r) => rect(s, r.x, r.y, 3.587, r.h, { fill: { color: RED } }));
	txt(s, 'Our Laboratory Services', { x: 2.127, y: 0.786, w: 9.08, h: 0.774, fontSize: 40, bold: true, align: 'center' });
	[
		{ x: 1.58, y: 2.864, title: 'Microbiology' },
		{ x: 5.168, y: 5.11, title: 'Physical Testing' },
		{ x: 8.755, y: 2.864, title: 'Chemical Testing' },
	].forEach((c) => {
		txt(s, c.title, { x: c.x + 0.199, y: c.y, w: 2.6, h: 0.37, fontFace: JK_SB, fontSize: 16, color: WHITE, align: 'center' });
		body(s, blurb, { x: c.x, y: c.y + 0.37, w: 2.998, h: 0.904, color: CREAM, align: 'center' });
	});
}

// 14 - Quality statement + 90% stat card
function slide14(pres) {
	const s = pres.addSlide();
	txt(s, [
		{ text: 'Committed to ', options: { breakLine: true } },
		{ text: 'Quality of Laboratory Products.' },
	], { x: 0.517, y: 1.763, w: 4.757, h: 1.616, fontSize: 30, bold: true });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed non eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis tortor nostrud exercitation ullamco laboris nisi ut aliquip enim.',
		{ x: 0.519, y: 3.528, w: 4.757, h: 1.182 });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed non eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis tortor nostrud.',
		{ x: 0.517, y: 4.832, w: 4.757, h: 0.904 });
	rect(s, 9.547, 1.014, 3.428, 2.59, { fill: { color: RED } });
	txt(s, '90%', { x: 10.632, y: 1.29, w: 1.259, h: 0.64, fontSize: 32, bold: true, color: WHITE, align: 'center', wrap: false });
	txt(s, 'Simple Text Here', { x: 9.962, y: 2.053, w: 2.6, h: 0.337, fontFace: JK_SB, fontSize: 14, color: WHITE, align: 'center' });
	body(s, 'PLACEHOLDER',
		{ x: 9.762, y: 2.423, w: 2.998, h: 0.904, color: CREAM, align: 'center' });
	txt(s, 'Biolatory.co', { x: 0.261, y: 0.304, w: 2.005, h: 0.339, fontFace: JK_SB, fontSize: 14 });
}

// 15 - Uncompromising quality + 1750+ card
function slide15(pres) {
	const s = pres.addSlide();
	body(s, 'Lorem dolor sed viverra ipsum nunc aliquet non eiusmod bibendum enim facilisis. Vestibulum lectus mauris ultrices eros minim cursus. Vel quam elementum pulvinar etiam non. Tortor id aliquet lectus proin nisl. ',
		{ x: 8.142, y: 5.287, w: 4.326, h: 1.179 });
	txt(s, 'Simple Text Here', { x: 8.142, y: 4.95, w: 2.842, h: 0.337, fontFace: JK_SB, fontSize: 14 });
	txt(s, 'Uncompromising Quality in Laboratory Analysis', { x: 0.789, y: 1.0, w: 6.794, h: 1.178, fontSize: 32, bold: true });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed tortor eiusmod tempor incididunt enim labore et dolore magna aliqua. Ut enim mollit minim veniam, quis nostrud exercitation ullamco laboris nisi sed aliquip deserunt commodo',
		{ x: 0.789, y: 2.274, w: 6.794, h: 0.904 });
	rect(s, 0.789, 3.57, 3.01, 3.16, { fill: { color: RED, transparency: 20 } });
	body(s, 'Lorem dolor sed viverra ipsum nunc aliquet bibendum tortor vestibulum lectus.',
		{ x: 1.063, y: 5.204, w: 2.461, h: 0.901, color: WHITE, align: 'center' });
	txt(s, 'Simple Text Here', { x: 1.063, y: 4.867, w: 2.461, h: 0.337, fontFace: JK_SB, fontSize: 14, color: WHITE, align: 'center' });
	txt(s, '1750+', { x: 1.471, y: 4.194, w: 1.645, h: 0.505, fontFace: JK_SB, fontSize: 24, bold: true, color: WHITE, align: 'center' });
}

// 16 - Project cards under a red hero band
function slide16(pres) {
	const s = pres.addSlide();
	rect(s, 0, 0, 13.333, 4.553, { fill: { color: RED, transparency: 15 } });
	[4.878, 0.977, 8.776].forEach((x, i) => {
		rect(s, x, 2.403, i === 2 ? 3.578 : 3.575, 4.274, { fill: { color: WHITE }, shadow: shadow('glowSoft') });
	});
	[
		{ x: 1.131, title: 'Our Project One' },
		{ x: 5.091, title: 'Our Project Two' },
		{ x: 9.052, title: 'Our Project Three' },
	].forEach((c) => {
		txt(s, c.title, { x: c.x, y: 4.88, w: 3.151, h: 0.337, fontFace: JK_SB, fontSize: 14 });
		body(s, 'Lorem ipsum dolor sit magna tempor consectetur adipiscing.', { x: c.x, y: 5.216, w: 3.151, h: 0.675, fontSize: 12 });
	});
	[3.846, 7.746, 11.648].forEach((x) => {
		rect(s, x, 6.153, 0.706, 0.525, { fill: { color: RED } });
		hline(s, x + 0.197, 6.415, 0.313, { line: { color: WHITE, width: 1.5, endArrowType: 'arrow' } });
	});
	txt(s, 'Our Last Project.', { x: 2.501, y: 0.493, w: 8.332, h: 0.707, fontSize: 36, bold: true, color: WHITE, align: 'center' });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed magna eiusmod tempor incididunt tortor labore dolore magna aliqua. Nisi scelerisque eu ultrices vitae auctor eu augue. ',
		{ x: 2.501, y: 1.29, w: 8.332, h: 0.627, color: WHITE, align: 'center' });
}

// 17 - Numbered steps stacked on the right
function slide17(pres) {
	const s = pres.addSlide();
	[
		{ y: 2.884, num: '02' },
		{ y: 4.921, num: '03' },
		{ y: 0.85, num: '01' },
	].forEach((c) => {
		rect(s, 7.414, c.y, 5.169, 1.731, { fill: { color: WHITE }, shadow: shadow('dropCool') });
		txt(s, 'Simple Text Here', { x: 8.637, y: c.y + 0.263, w: 3.111, h: 0.337, fontFace: JK_SB, fontSize: 14 });
		body(s, 'Lorem dolor sed viverra ipsum nunc aliquet minim bibendum enim facilisis. Vestibulum enim lectus mauris ultrices eros cursus. ',
			{ x: 8.637, y: c.y + 0.568, w: 3.705, h: 0.901 });
		txt(s, c.num, {
			shape: S.rect, x: 7.656, y: c.y + 0.349, w: 0.793, h: 0.793,
			fill: { color: RED }, fontFace: JK_SB, fontSize: 24, color: WHITE, align: 'center', valign: 'middle',
		});
	});
	txt(s, 'Integrity and Innovation in Medical Diagnostics.', { x: 0.75, y: 4.173, w: 5.917, h: 1.178, fontSize: 32, bold: true });
	body(s, 'Lorem dolor sed viverra ipsum nunc aliquet bibendum enim facilisis. Vestibulum lectus mauris ultrices eros in cursus. Vel quam elementum pulvinar etiam tortor tortor id aliquet lectus proin.',
		{ x: 0.75, y: 5.534, w: 5.917, h: 0.901 });
}

// 18 - Donut gauge + numbered notes
function slide18(pres) {
	const s = pres.addSlide();
	rect(s, 0, 0, 2.534, 7.5, { fill: { color: RED } });
	rect(s, 4.503, 4.053, 2.702, 2.547, { fill: { color: WHITE }, shadow: shadow('haloSoft') });
	txt(s, 'Your Title Here ', { x: 4.99, y: 5.921, w: 1.729, h: 0.337, fontFace: JK_SB, fontSize: 14, bold: true, align: 'center' });
	donut(s, pres, 5.165, 4.337, 1.362, RED);
	txt(s, '80%', { x: 5.126, y: 4.837, w: 1.456, h: 0.404, fontFace: JK_SB, align: 'center' });
	txt(s, 'Delivering Quality Results You Can Trust', { x: 8.115, y: 1.407, w: 4.484, h: 1.043, fontSize: 28, bold: true });
	body(s, 'Turpis egestas pretium aenean pharetra magna ac placerat. Cras sed felis eget velit aliquet sagittis id consectetur purus. Habitant morbi tristique senectus et netus et. Blandit turpis cursus in hac habitasse lorem ipsum',
		{ x: 8.115, y: 2.611, w: 4.484, h: 1.179 });
	[
		{ y: 3.984, num: '01', circleX: 8.15, circleY: 4.026, textX: 8.221, textW: 0.447 },
		{ y: 5.18, num: '02', circleX: 8.155, circleY: 5.222, textX: 8.207, textW: 0.485 },
	].forEach((r) => {
		body(s, 'PLACEHOLDER',
			{ x: 8.839, y: r.y + 0.286, w: 3.204, h: 0.627 });
		txt(s, 'Simple Text Here', { x: 8.839, y: r.y, w: 2.702, h: 0.337, fontFace: JK_SB, fontSize: 14 });
		oval(s, r.circleX, r.circleY, 0.59, 0.59, { fill: { color: RED }, shadow: shadow('haloWide') });
		txt(s, r.num, { x: r.textX, y: r.circleY + 0.121, w: r.textW, h: 0.348, fontFace: JK_SB, fontSize: 12, color: WHITE, align: 'center' });
	});
}

// 19 - Red hero card with two captions below
function slide19(pres) {
	const s = pres.addSlide();
	rect(s, 0.941, 0.728, 5.725, 3.661, { fill: { color: RED }, shadow: shadow('dropCool') });
	const caption = 'PLACEHOLDER';
	body(s, caption, { x: 2.992, y: 5.539, w: 3.262, h: 0.905, align: 'justify' });
	txt(s, 'Simple Text Here', { x: 2.992, y: 5.186, w: 2.587, h: 0.337, fontFace: JK_SB, fontSize: 14 });
	body(s, caption, { x: 8.706, y: 5.539, w: 3.237, h: 0.904, align: 'justify' });
	txt(s, 'Simple Text Here', { x: 8.706, y: 5.186, w: 2.159, h: 0.337, fontFace: JK_SB, fontSize: 14 });
	txt(s, 'The Laboratory For Analytical Chemistry.',
		{ x: 1.291, y: 1.316, w: 5.027, h: 1.111, fontSize: 30, bold: true, color: WHITE, align: 'center' });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed non eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis tortor nostrud exercitation ullamco laboris nisi ut aliquip enim.',
		{ x: 1.291, y: 2.619, w: 5.027, h: 1.182, color: WHITE, align: 'center' });
}

// 20 - Gallery title / caption (image frames come from the layout)
function slide20(pres) {
	const s = pres.addSlide();
	txt(s, 'Our Amazing Gallery', { x: 3.319, y: 0.806, w: 6.695, h: 0.774, fontSize: 40, bold: true, align: 'center' });
	body(s, 'Lorem dolor sed viverra ipsum nunc aliquet bibendum enim facilisis. Vestibulum lectus mauris ultrices eros in cursus. Vel quam elementum pulvinar etiam non. Tortor id aliquet lectus proin nibh nisl. Euismod lacinia at quis risus sed vulputate odio. Euismod elementum nisi quis eleifend quam. ',
		{ x: 2.077, y: 5.926, w: 9.179, h: 0.901, align: 'center' });
}

// 21 - Phone mock-ups over a diagonal red wedge
function slide21(pres) {
	const s = pres.addSlide();
	redWedge(s, 0, false);
	// two tilted phone mock-ups: white screen inside a dark bezel
	[{ x: 0.297, y: -0.983 }, { x: 2.324, y: 2.719 }].forEach((p) => {
		imagePlaceholder(s, p.x, p.y, 3.214, 6.445,
			{ rotate: 35.3, fill: WHITE, line: '2B2B2B', lineWidth: 6, rounded: 0.35 });
	});
	txt(s, 'Providing Insights For Effective Treatments.', { x: 7.536, y: 2.113, w: 5.209, h: 1.178, fontSize: 32, bold: true });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit quis sed tortor eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim mollit minim veniam, quis nostrud exercitation ullamco laboris aliquip deserunt commodo consequat duis aute irure dolor.',
		{ x: 7.536, y: 3.391, w: 5.209, h: 1.182 });
	button(s, 7.642, 4.886, 1.757, 0.501, 'Learn More');
}

// 22 - Laptop mock-up + two donut gauges
function slide22(pres) {
	const s = pres.addSlide();
	redWedge(s, 7.685, true);
	// laptop mock-up: white screen in a dark bezel, running off the right edge
	imagePlaceholder(s, 6.376, 1.016, 9.448, 5.469, { fill: WHITE, line: '2B2B2B', lineWidth: 6 });
	txt(s, 'Science is The Foundation of Our Prosperity.', { x: 0.827, y: 1.1, w: 4.638, h: 1.818, fontSize: 34, bold: true });
	body(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit quis sed tortor eiusmod tempor incididunt ut labore et dolore non magna aliqua. Ut enim mollit minim veniam',
		{ x: 0.827, y: 3.073, w: 4.638, h: 0.904 });
	[
		{ x: 0.918, fill: RED, arc: AMBER, label: '80%', labelColor: WHITE, titleColor: WHITE },
		{ x: 3.345, fill: WHITE, arc: RED, label: '90%', labelColor: BLACK, titleColor: BLACK },
	].forEach((g) => {
		rect(s, g.x, 4.357, 2.064, 2.043, { fill: { color: g.fill }, shadow: shadow('dropCool') });
		txt(s, 'Your Title Here ', { x: g.x + 0.167, y: 5.815, w: 1.729, h: 0.337, fontFace: JK_SB, fontSize: 14, color: g.titleColor, align: 'center' });
		donut(s, pres, g.x + 0.458, 4.604, 1.064, g.arc);
		txt(s, g.label, { x: g.x + 0.428, y: 4.954, w: 1.138, h: 0.404, fontFace: JK_SB, bold: true, color: g.labelColor, align: 'center' });
	});
}

// 23 - Contact card on a red background
function slide23(pres) {
	const s = pres.addSlide();
	s.background = { color: RED };
	rect(s, 6.667, 0.803, 5.604, 6.697, { fill: { color: WHITE }, shadow: shadow('dropCool') });
	body(s, 'Laoreet suspendisse interdum consectetur libero id. Sed sed risus pretium quam vulputate dignissim suspendisse risus in hendrerit gravida rutrum quisque non. ',
		{ x: 7.214, y: 2.036, w: 4.509, h: 0.901, color: GRAY_MID });
	txt(s, [
		{ text: 'East 68th Street, ', options: { breakLine: true } },
		{ text: 'New York, NY 10065' },
	], { x: 8.254, y: 3.218, w: 3.4, h: 0.64, fontFace: JK_SB, fontSize: 16 });
	txt(s, 'www.yourcompany.com', { x: 8.254, y: 4.218, w: 3.058, h: 0.37, fontFace: JK_SB, fontSize: 16 });
	txt(s, '+123 456 7890', { x: 8.254, y: 5.12, w: 2.66, h: 0.37, fontFace: JK_SB, fontSize: 16 });
	iconMark(s, 7.246, 3.186, 0.656, 0.656, RED);
	iconMark(s, 7.301, 4.13, 0.547, 0.547, RED);
	iconMark(s, 7.214, 4.95, 0.72, 0.72, RED);
	[4.052, 4.848, 5.681].forEach((y) => hline(s, 8.367, y, 3.287, { line: { color: GRAY_LINE, width: 0.75 } }));
	txt(s, 'Simple Text Here', { x: 7.214, y: 1.707, w: 3.357, h: 0.337, fontFace: JK_SB, fontSize: 14 });
	button(s, 7.301, 6.103, 1.662, 0.494, 'Contact Now', { fill: BLACK, fontFace: INTER_MED, fontSize: 12, shadow: shadow('lift') });
	txt(s, 'If Need Any Info Please Contact Us!', { x: 0.701, y: 4.857, w: 5.604, h: 1.447, fontSize: 40, bold: true, color: WHITE });
}

// 24 - Thank you
function slide24(pres) {
	const s = pres.addSlide();
	rect(s, 0, 2.615, 7.671, 2.271, { fill: { color: WHITE }, shadow: shadow('liftStrong') });
	txt(s, 'THANK YOU', { x: 0.592, y: 3.094, w: 6.487, h: 1.313, fontSize: 72, bold: true, align: 'center' });
	txt(s, 'LABORATORY & SCIENCE PRESENTATION', {
		shape: S.rect, x: 2.935, y: 4.885, w: 5.628, h: 0.513,
		fill: { color: RED }, shadow: shadow('liftStrong'),
		fontSize: 12, color: WHITE, charSpacing: 1, align: 'center', valign: 'middle',
	});
}

/* ================================================================= assembly */
function build() {
	const pres = new PptxGenJS();
	S = pres.ShapeType;
	pres.defineLayout({ name: 'WIDE', width: 40 / 3, height: 7.5 }); // 12192000 x 6858000 EMU
	pres.layout = 'WIDE';
	pres.author = 'Biolatory';
	pres.title = 'Biolatory - Laboratory & Science Presentation';

	[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
		slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16,
		slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24,
	].forEach((fn) => fn(pres));

	return pres.writeFile({ fileName: path.join(__dirname, '07ac80b5-049a-46af-8164-e730ddb7c567_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
