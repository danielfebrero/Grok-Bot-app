/*
 * Recreation of "Understanding the Property Investment Landscape" (15 slides, 13.333" x 7.5")
 * with pptxgenjs only.  Raster artwork in the source deck (icon PNGs, the tablet photo)
 * is replaced by native pptxgenjs shapes drawn in code.
 *
 *   node 01bb4b9d-f971-45cc-b9b0-f98235925c84_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ theme */

const PURPLE = '7D3AFC';   // theme accent1
const DARK   = '262626';   // theme accent2
const CREAM  = 'FEFAD5';   // theme accent3 (also the master background)
const WHITE  = 'FFFFFF';
const BLACK  = '000000';
const GRAY   = 'E7E6E6';   // theme bg2 - used for drop shadows / dashed rings

const HEAD = 'Montserrat SemiBold';   // theme major latin font
const BODY = 'Montserrat';            // theme minor latin font

const NBSP = '\u00A0';

/* --------------------------------------------------------------- helpers */

const NO_LINE = { type: 'none' };

/** soft drop shadow, mirroring the deck's `outerShdw` effects */
const shadow = (color, opacity, blur, offset = 0, angle = 50) =>
	({ type: 'outer', color, opacity, blur, offset, angle, rotateWithShape: false });

/** solid shape */
function box(slide, shape, x, y, w, h, color, extra = {}) {
	slide.addShape(shape, { x, y, w, h, fill: { color }, line: NO_LINE, ...extra });
}

/** hollow ellipse - the hand-drawn "circled word" motif used on every title */
function ring(slide, x, y, w, h, color, width = 1) {
	slide.addShape('ellipse', { x, y, w, h, fill: { type: 'none' }, line: { color, width } });
}

/* PowerPoint's default text-box insets, in points: left/right 0.1", top/bottom 0.05".
 * pptxgenjs takes them as [left, right, bottom, top]. */
const INSET = [7.2, 7.2, 3.6, 3.6];

/** heading text (major font, top anchored like the original auto-fit text boxes) */
function heading(slide, text, x, y, w, h, fontSize, color, extra = {}) {
	slide.addText(text, {
		x, y, w, h, fontFace: HEAD, fontSize, color,
		valign: 'top', isTextBox: true, margin: INSET, ...extra,
	});
}

/** body copy (minor font, 12pt / 130% leading is the deck default) */
function para(slide, text, x, y, w, h, extra = {}) {
	slide.addText(text, {
		x, y, w, h, fontFace: BODY, fontSize: 12, color: DARK, lineSpacingMultiple: 1.3,
		valign: 'top', isTextBox: true, margin: INSET, ...extra,
	});
}

/** free-form polygon: `pts` are [x, y] fractions of the shape box, "c" entries are cubic curves */
function poly(slide, x, y, w, h, pts, color, extra = {}) {
	const points = pts.map((p) => {
		if (p.length === 6) {
			return { x: w * p[4], y: h * p[5],
				curve: { type: 'cubic', x1: w * p[0], y1: h * p[1], x2: w * p[2], y2: h * p[3] } };
		}
		return { x: w * p[0], y: h * p[1] };
	});
	points.push({ close: true });
	slide.addShape('custGeom', { x, y, w, h, points, fill: { color }, line: NO_LINE, ...extra });
}

/* The four-stroke "sparkle" accent that decorates most slides.  In the source deck it is a
 * group of four rotated PNG strips; here the strips are drawn as rotated rounded bars.
 * Geometry is expressed as fractions of the group box so any group size/rotation works. */
const SPARKLE_BARS = [
	{ cx: 0.139, cy: 0.639, len: 0.754, thick: 0.209, angle: -38.2 },
	{ cx: 0.394, cy: 0.412, len: 0.888, thick: 0.246, angle: -7.8 },
	{ cx: 0.606, cy: 0.412, len: 0.888, thick: 0.246, angle: 7.8 },
	{ cx: 0.861, cy: 0.639, len: 0.754, thick: 0.209, angle: 38.2 },
];

function sparkle(slide, x, y, w, h, color, rotate = 0) {
	const rad = (rotate * Math.PI) / 180, cos = Math.cos(rad), sin = Math.sin(rad);
	SPARKLE_BARS.forEach((b) => {
		const dx = (b.cx - 0.5) * w, dy = (b.cy - 0.5) * h;
		const px = x + w / 2 + dx * cos - dy * sin;
		const py = y + h / 2 + dx * sin + dy * cos;
		const bw = b.thick * h, bh = b.len * h;
		box(slide, 'roundRect', px - bw / 2, py - bh / 2, bw, bh, color,
			{ rectRadius: 0.02, rotate: b.angle + rotate });
	});
}

/* Jigsaw piece: a diamond whose four edges carry a round knob - a "tab" in the piece
 * colour bulging outwards, or a "blank" in the background colour biting in.
 * Edge keys are the compass directions of the diamond's four sides. */
const EDGE_NORMALS = { ne: [1, -1], se: [1, 1], sw: [-1, 1], nw: [-1, -1] };

function puzzlePiece(slide, x, y, w, h, color, tabs, bg = CREAM, knob = 0.0743, seat = 0.25) {
	box(slide, 'diamond', x, y, w, h, color);
	Object.keys(tabs).forEach((edge) => {
		const [sx, sy] = EDGE_NORMALS[edge];
		const len = Math.hypot(h, w);
		const nx = (sx * h) / len, ny = (sy * w) / len;             // outward unit normal
		const push = (tabs[edge] === 'tab' ? 0.5 : -0.6) * knob * w;  // tabs sit proud, blanks bite in
		const cx = x + w / 2 + sx * seat * w + nx * push;
		const cy = y + h / 2 + sy * seat * h + ny * push;
		box(slide, 'ellipse', cx - knob * w, cy - knob * h, 2 * knob * w, 2 * knob * h,
			tabs[edge] === 'tab' ? color : bg);
	});
}

/* Icon placeholders standing in for the small PNG/SVG glyphs of the source deck. */
function icon(slide, kind, x, y, size, color) {
	const cx = x + size / 2, cy = y + size / 2;
	if (kind === 'arrow') {
		slide.addShape('line', {
			x: x + 0.06 * size, y: cy, w: size * 0.88, h: 0,
			line: { color, width: Math.max(1, size * 4), endArrowType: 'triangle' },
		});
	} else if (kind === 'coins') {
		box(slide, 'can', cx - size * 0.42, cy - size * 0.40, size * 0.62, size * 0.80, color);
		box(slide, 'ellipse', cx - size * 0.02, cy + size * 0.02, size * 0.42, size * 0.42, color);
	} else if (kind === 'target') {
		box(slide, 'donut', x, y, size, size, color);
		box(slide, 'ellipse', cx - size * 0.16, cy - size * 0.16, size * 0.32, size * 0.32, color);
	} else if (kind === 'chart') {
		[0.30, 0.55, 0.80].forEach((f, i) => {
			box(slide, 'rect', x + size * (0.12 + i * 0.28), y + size * (1 - f) + size * 0.05,
				size * 0.2, size * (f - 0.05), color);
		});
	} else if (kind === 'code') {
		box(slide, 'roundRect', x, y + size * 0.10, size, size * 0.62, color, { rectRadius: 0.02 });
		box(slide, 'rect', cx - size * 0.22, y + size * 0.74, size * 0.44, size * 0.10, color);
		box(slide, 'ellipse', x + size * 0.58, y + size * 0.56, size * 0.44, size * 0.44, color);
	} else if (kind === 'star') {
		box(slide, 'star4', x, y, size, size, color, { rotate: 15 });
	}
}

/* ---------------------------------------------------------------- slides */

/* 1 - title */
function slide01(pptx) {
	const s = pptx.addSlide();
	s.background = { color: PURPLE };
	heading(s, 'Understanding the Property Investment Landscape',
		0.68, 2.336, 11.973, 2.827, 60, CREAM, { align: 'center', lineSpacingMultiple: 0.9 });
	ring(s, 6.112, 3.286, 4.99, 0.929, CREAM, 1.5);
	sparkle(s, 5.934, 1.287, 1.466, 0.646, CREAM);
	// dome / half-round sitting on the bottom edge
	poly(s, 4.966, 5.799, 3.401, 1.701, [
		[0.5, 0], [0.776, 0, 1, 0.448, 1, 1], [0, 1], [0, 0.448, 0.224, 0, 0.5, 0],
	], CREAM);
}

/* 2 - four "step" circles */
const STEPS = [
	{ x: 0.881, label: 'Step One',   fill: CREAM,  ink: BLACK },
	{ x: 3.812, label: 'Step Two',   fill: PURPLE, ink: CREAM },
	{ x: 6.743, label: 'Step Three', fill: CREAM,  ink: BLACK },
	{ x: 9.673, label: 'Step Four',  fill: CREAM,  ink: BLACK },
];

function slide02(pptx) {
	const s = pptx.addSlide();
	s.background = { color: DARK };
	heading(s, 'The Fundamentals of Real Estate Investment', 0.67, 0.9, 7.279, 1.447, 40, CREAM);
	ring(s, 4.086, 1.601, 3.373, 0.724, CREAM);
	para(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium,' +
		NBSP + NBSP + 'totam rem', 9.673, 1.085, 2.99, 1.127, { color: WHITE });

	STEPS.forEach((st) => {
		box(s, 'ellipse', st.x, 3.263, 2.779, 2.779, st.fill,
			st.fill === PURPLE ? { shadow: shadow(PURPLE, 0.15, 25) } : {});
		heading(s, st.label, st.x + 0.594, 3.811, 1.591, 0.37, 16, st.ink, { align: 'center' });
		para(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ',
			st.x + 0.274, 4.184, 2.231, 0.865, { align: 'center', color: st.ink });
		icon(s, 'arrow', st.x + 1.131, 5.18, 0.518, st.ink);
	});

	sparkle(s, 0.592, 3.14, 0.932, 0.411, CREAM, 321.4);
	sparkle(s, 8.824, 5.729, 0.932, 0.411, CREAM, 135);
}

/* 3 - single purple card (the right half holds an empty picture placeholder) */
function slide03(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };
	box(s, 'roundRect', 0.594, 3.75, 4.016, 3.226, PURPLE,
		{ rectRadius: 0.184, shadow: shadow(PURPLE, 0.15, 25) });
	heading(s, 'From Land to Legacy Property  Investment Secures Financial Freedom for Modern Property Investors',
		0.901, 4.277, 3.354, 2.249, 20, CREAM, { lineSpacingMultiple: 1.3 });
	sparkle(s, 0.365, 3.991, 0.932, 0.411, CREAM, 321.4);
	ring(s, 0.955, 5.188, 2.617, 0.462, CREAM);
}

/* 4 - "Unlocking Opportunities" */
function slide04(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };
	heading(s, 'Unlocking Opportunities Future of Property Investment', 0.467, 0.407, 8.495, 2.121, 40, DARK);
	ring(s, 2.602, 1.105, 3.373, 0.724, PURPLE);
	// purple panel bleeding off the right edge, left corners rounded
	poly(s, 10.708, 3.75, 2.626, 3.226, [
		[0.07, 0], [1, 0], [1, 1], [0.07, 1],
		[0.031, 1, 0, 0.974, 0, 0.943], [0, 0.057], [0, 0.026, 0.031, 0, 0.07, 0],
	], PURPLE, { shadow: shadow(PURPLE, 0.15, 25) });
	para(s, 'natus error sit voluptatem accusantium doloremque laudantium,' + NBSP + NBSP + 'totam rem',
		10.425, 1.035, 2.441, 0.865);
	sparkle(s, 0.026, 3.356, 0.932, 0.411, PURPLE, 321.4);
	// cream quarter-disc badge in the bottom right corner
	poly(s, 12.43, 6.073, 0.903, 0.903, [
		[1, 0], [1, 1], [0, 1], [0, 0.448, 0.448, 0, 1, 0],
	], CREAM);
	heading(s, 'Navigating Confidence Market Invest', 10.974, 4.031, 2.182, 1.374, 20, CREAM,
		{ lineSpacingMultiple: 1.3 });
	heading(s, '15+', 12.672, 6.432, 0.631, 0.438, 20, DARK, { align: 'center' });
}

/* 5 - "Lucrative Property Market Invesment" */
function slide05(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };
	box(s, 'roundRect', 9.283, 0.524, 3.456, 3.226, DARK,
		{ rectRadius: 0.184, shadow: shadow(PURPLE, 0.15, 25) });
	heading(s, 'Precentage ', 9.581, 0.808, 1.778, 0.404, 18, WHITE, { wrap: false });
	heading(s, '90%', 9.581, 2.238, 2.77, 1.313, 72, WHITE);
	sparkle(s, 9.207, 2.158, 0.932, 0.411, PURPLE, 321.4);

	box(s, 'roundRect', 0.594, 3.75, 3.083, 3.226, PURPLE,
		{ rectRadius: 0.176, shadow: shadow(PURPLE, 0.15, 25) });
	heading(s, 'Navigating the Real Estate with Confidence and Clarity Market',
		0.864, 4.011, 2.547, 1.811, 20, CREAM, { lineSpacingMultiple: 1.3 });
	heading(s, 'Read More', 0.864, 6.391, 1.135, 0.303, 12, WHITE, { wrap: false });
	icon(s, 'arrow', 2.724, 6.284, 0.518, CREAM);

	heading(s, 'Lucrative Property Market Invesment', 4.192, 4.111, 5.586, 1.447, 40, DARK);
	ring(s, 6.269, 4.835, 3.373, 0.724, PURPLE);
	para(s, 'natus error sit voluptatem accusantium doloremque laudantium,' + NBSP + NBSP + 'totam rem',
		10.47, 5.822, 2.441, 0.865);
}

/* 6 - team strip */
const TEAM = [
	{ x: 1.507, y: 5.873, fill: PURPLE, ink: CREAM },
	{ x: 5.698, y: 5.870, fill: CREAM,  ink: DARK  },
	{ x: 9.870, y: 5.870, fill: PURPLE, ink: CREAM },
];

function slide06(pptx) {
	const s = pptx.addSlide();
	s.background = { color: DARK };
	heading(s, 'Strategic Real Estate Investments Turning Building Wealth Through Smart and Assets into Long-Term Financial Gains',
		1.545, 0.709, 10.244, 1.717, 32, CREAM, { align: 'center' });
	ring(s, 7.447, 1.755, 3.666, 0.724, CREAM);
	sparkle(s, 4.952, 3.299, 0.932, 0.411, CREAM, 321.4);

	TEAM.forEach((c) => {
		box(s, 'roundRect', c.x, c.y, 2.666, 0.764, c.fill,
			{ rectRadius: 0.166, shadow: shadow(c.fill === CREAM ? DARK : PURPLE, 0.15, 25) });
		heading(s, 'Alex Fradison', c.x + 0.081, c.y + 0.057, 1.757, 0.37, 16, c.ink, { wrap: false });
		para(s, 'Job Position', c.x + 0.081, c.y + 0.407, 1.233, 0.303,
			{ color: c.ink, lineSpacingMultiple: 1 });
		icon(s, 'arrow', c.x + 2.248, c.y + 0.428, 0.23, c.ink);
	});
}

/* 7 - tablet mock-up (the source deck uses a photo; drawn natively here) */
function slide07(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };
	box(s, 'roundRect', 7.718, 0.808, 4.528, 5.919, '1E1E1E', { rectRadius: 0.2 });   // [image] tablet frame
	box(s, 'rect', 7.863, 1.048, 4.238, 5.519, BLACK);
	heading(s, 'PLACEHOLDER',
		0.413, 0.43, 5.202, 3.467, 40, DARK);
	ring(s, 0.29, 2.459, 2.793, 0.724, PURPLE);
	sparkle(s, 7.275, 0.486, 0.932, 0.411, PURPLE, 321.4);
	sparkle(s, 11.883, 6.522, 0.932, 0.411, DARK, 135);
	para(s, 'lorem ipsum legit dolor sit amet, consectetuer et cum adipiscing elit. aenean commodo ligula eget dolor. aenean massa. cum sociis natoque pentibus cum adipiscing elit. ',
		0.413, 5.815, 5.177, 0.865, { align: 'justify' });
}

/* 8 - two feature cards + four numbered chevrons */
const FEATURES = [
	{ y: 2.651, fill: PURPLE, ink: WHITE, glyph: 'target', glyphColor: PURPLE },
	{ y: 4.205, fill: DARK,   ink: CREAM, glyph: 'star',   glyphColor: '404040' },
];
const CHEVRONS = [
	{ y: 0.969, fill: PURPLE, mirror: false, num: '01', numColor: PURPLE, numX: 8.355, numW: 0.710 },
	{ y: 2.471, fill: DARK,   mirror: true,  num: '02', numColor: DARK,   numX: 10.762, numW: 0.817 },
	{ y: 3.974, fill: PURPLE, mirror: false, num: '03', numColor: PURPLE, numX: 8.294, numW: 0.831 },
	{ y: 5.477, fill: DARK,   mirror: true,  num: '04', numColor: DARK,   numX: 10.740, numW: 0.886 },
];

function slide08(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };
	heading(s, 'Invest Smarter Grow Property', 1.034, 0.687, 4.483, 1.447, 40, DARK);
	ring(s, 2.661, 1.4, 2.793, 0.724, PURPLE);

	FEATURES.forEach((f, i) => {
		box(s, 'roundRect', 1.12, f.y, 4.335, 1.382, f.fill,
			{ rectRadius: 0.128, shadow: shadow('0D0D0D', 0.08, 32, 10, 50) });
		heading(s, 'Your text goes here', 2.178, f.y + 0.258, 2.77, 0.326, 16, f.ink,
			{ lineSpacingMultiple: 0.8 });
		para(s, 'A small river named Duden flows by their place and supplies it',
			2.178, f.y + 0.551, 3.117, 0.605, { color: f.ink });
		box(s, 'roundRect', 1.42, f.y + 0.368, 0.647, 0.65, WHITE,
			{ rectRadius: 0.209, shadow: shadow(BLACK, 0.18, 45, 10, 50) });
		icon(s, f.glyph, 1.61, f.y + 0.55, 0.27, f.glyphColor);
	});

	para(s, 'lorem ipsum legit dolor sit amet, consectetuer et cum adipiscing elit. aenean commodo ligula',
		1.035, 6.14, 4.335, 0.602, { align: 'justify' });

	// vertical light-gray rail behind the chevrons
	box(s, 'rect', 6.212, 3.474, 7.5, 0.551, 'E9E9E9', { rotate: 90 });

	CHEVRONS.forEach((c) => {
		if (c.mirror) {
			box(s, 'homePlate', 9.154, c.y + 0.138, 2.384, 1.018, c.fill,
				{ flipH: true, shadow: shadow(PURPLE, 0.2, 12, 10, 50) });
			box(s, 'chevron', 10.465, c.y, 1.749, 1.289, WHITE,
				{ flipH: true, shadow: shadow(BLACK, 0.1, 25, 10, 45) });
		} else {
			box(s, 'homePlate', 8.369, c.y + 0.138, 2.384, 1.018, c.fill,
				{ shadow: shadow(PURPLE, 0.2, 12, 10, 50) });
			box(s, 'chevron', 7.682, c.y, 1.749, 1.289, WHITE,
				{ shadow: shadow(BLACK, 0.1, 25, 10, 45) });
		}
		heading(s, c.num, c.numX, c.y + 0.281, c.numW, 0.707, 36, c.numColor,
			{ align: 'center', wrap: false });
	});
}

/* 9 - S-curve timeline: six alternating nodes joined by half-circle arcs */
const TIMELINE = [
	{ arcX: 1.607,  arcY: 3.604, dotX: 1.890,  dotY: 3.887, up: false, color: PURPLE, year: '2025', yearX: 1.807,  yearY: 5.570, textX: 1.487,  textY: 6.001 },
	{ arcX: 3.301,  arcY: 3.636, dotX: 3.584,  dotY: 3.919, up: true,  color: DARK,   year: '2026', yearX: 3.499,  yearY: 2.350, textX: 3.180,  textY: 2.781 },
	{ arcX: 4.994,  arcY: 3.614, dotX: 5.278,  dotY: 3.898, up: false, color: PURPLE, year: '2027', yearX: 5.182,  yearY: 5.570, textX: 4.862,  textY: 6.001 },
	{ arcX: 6.667,  arcY: 3.636, dotX: 6.950,  dotY: 3.919, up: true,  color: DARK,   year: '2028', yearX: 6.867,  yearY: 2.346, textX: 6.547,  textY: 2.777 },
	{ arcX: 8.360,  arcY: 3.604, dotX: 8.644,  dotY: 3.887, up: false, color: PURPLE, year: '2029', yearX: 8.557,  yearY: 5.570, textX: 8.237,  textY: 6.001 },
	{ arcX: 10.033, arcY: 3.636, dotX: 10.316, dotY: 3.919, up: true,  color: DARK,   year: '2030', yearX: 10.267, yearY: 2.346, textX: 9.948,  textY: 2.777 },
];

function slide09(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };
	heading(s, 'PLACEHOLDER',
		1.545, 0.503, 10.244, 1.178, 32, DARK, { align: 'center' });
	ring(s, 5.446, 1.037, 3.566, 0.618, PURPLE);

	TIMELINE.forEach((n) => {
		s.addShape('arc', {
			x: n.arcX, y: n.arcY, w: 1.683, h: 1.683,
			angleRange: [270, 88], rotate: n.up ? 270 : 90, flipV: n.up,
			fill: { type: 'none' },
			line: { color: n.color, width: 3, beginArrowType: 'triangle' },
		});
		box(s, 'ellipse', n.dotX, n.dotY, 1.116, 1.116, WHITE, { shadow: shadow(BLACK, 0.2, 25, 10, 45) });
		box(s, 'ellipse', n.dotX + 0.168, n.dotY + 0.168, 0.781, 0.781, n.color);
		heading(s, n.year, n.yearX, n.yearY, 1.315, 0.438, 20, n.year === '2028' ? BLACK : n.color,
			{ align: 'center' });
		para(s, 'Lorem ipsum dolor sit amet, consectet', n.textX, n.textY, 1.922, 0.608,
			{ align: 'center' });
	});
}

/* 10 - isometric four-piece puzzle + four percentage read-outs */
const PUZZLE = [
	{ x: 2.425, y: 1.456, w: 2.778, h: 1.610, color: PURPLE, cirX: 3.504, cirY: 2.100, glyph: 'coins',  glyphColor: PURPLE,
	  tabs: { nw: 'blank', ne: 'blank', se: 'tab', sw: 'tab' } },
	{ x: 0.690, y: 2.516, w: 2.776, h: 1.610, color: DARK,   cirX: 1.832, cirY: 2.972, glyph: 'code',   glyphColor: DARK,
	  tabs: { nw: 'blank', ne: 'blank', se: 'tab', sw: 'blank' } },
	{ x: 4.112, y: 2.478, w: 2.778, h: 1.610, color: DARK,   cirX: 5.117, cirY: 2.972, glyph: 'chart',  glyphColor: DARK,
	  tabs: { nw: 'blank', ne: 'blank', se: 'blank', sw: 'tab' } },
	{ x: 2.414, y: 3.553, w: 2.776, h: 1.613, color: PURPLE, cirX: 3.472, cirY: 4.177, glyph: 'target', glyphColor: PURPLE,
	  tabs: { nw: 'blank', ne: 'blank', se: 'tab', sw: 'tab' } },
];
const STATS = [
	{ x: 7.733, y: 3.184, pct: '68%', color: PURPLE, labelX: 7.765 },
	{ x: 10.001, y: 3.184, pct: '72%', color: DARK,   labelX: 10.039 },
	{ x: 7.733, y: 5.060, pct: '65%', color: DARK,   labelX: 7.765 },
	{ x: 9.931, y: 5.060, pct: '80%', color: PURPLE, labelX: 10.039 },
];

function slide10(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };

	PUZZLE.forEach((p) => puzzlePiece(s, p.x - 0.061, p.y + 0.416, p.w, p.h, GRAY, p.tabs, CREAM));  // cast shadow
	PUZZLE.forEach((p) => puzzlePiece(s, p.x, p.y, p.w, p.h, p.color, p.tabs, CREAM));
	PUZZLE.forEach((p) => {
		box(s, 'ellipse', p.cirX, p.cirY, 0.637, 0.637, WHITE);
		icon(s, p.glyph, p.cirX + 0.155, p.cirY + 0.155, 0.327, p.glyphColor);
	});

	heading(s, 'Invest Smarter Grow Property', 7.733, 0.83, 4.483, 1.447, 40, DARK);
	ring(s, 9.36, 1.544, 2.793, 0.724, PURPLE);

	STATS.forEach((t) => {
		heading(s, t.pct, t.x, t.y, 2.198, 0.669, 32, t.color, { lineSpacingMultiple: 1.1 });
		heading(s, 'Your Project', t.labelX, t.y + 0.583, 2.179, 0.446, 16, BLACK, { lineSpacingMultiple: 1.4 });
		para(s, 'Lorem ipsum dolor sit amet, consectetur', t.labelX, t.y + 0.983, 2.179, 0.64,
			{ color: BLACK, lineSpacingMultiple: 1.4 });
	});

	para(s, 'lorem ipsum legit dolor sit amet, consectetuer et cum adipiscing elit. aenean commodo ligula adispacing mountain legit cum et',
		0.713, 6.076, 5.953, 0.602, { align: 'justify' });
}

/* 11 - four circular "year" badges.  The badge body is a circle with a pill-shaped
 * tab at the lower left, reproduced here as a custom geometry. */
const BADGE_BLOB = [
	[0.582, 0],
	[0.813, 0, 1, 0.224, 1, 0.5],
	[1, 0.759, 0.835, 0.972, 0.625, 0.997],
	[0.582, 1.0],
	[0.091, 0.999],
	[0.053, 0.999, 0.021, 0.972, 0.007, 0.933],
	[0, 0.891],
	[0.007, 0.848],
	[0.021, 0.809, 0.053, 0.782, 0.091, 0.782],
	[0.237, 0.782],
	[0.235, 0.780],
	[0.190, 0.700, 0.163, 0.604, 0.163, 0.5],
	[0.163, 0.224, 0.351, 0, 0.582, 0],
];
const BADGES = [
	{ numX: 2.810,  numY: 2.280, numW: 1.035, numH: 1.106, num: '01', numColor: PURPLE, numAlpha: 80,
	  ringX: 1.163, ringY: 2.634, ringColor: 'CBB0FE', blobX: 1.073, blobY: 2.878, color: PURPLE,
	  year: '2024', yearX: 1.087, yearY: 4.218, yearW: 0.751, capX: 1.401, capY: 3.279, txtX: 1.505, txtY: 3.545 },
	{ numX: 5.317,  numY: 3.703, numW: 1.182, numH: 1.076, num: '02', numColor: DARK, numAlpha: 80,
	  ringX: 3.670, ringY: 4.057, ringColor: 'A8A8A8', blobX: 3.580, blobY: 4.301, color: DARK,
	  year: '2025', yearX: 3.591, yearY: 5.641, yearW: 0.747, capX: 3.897, capY: 4.725, txtX: 4.000, txtY: 4.991 },
	{ numX: 8.533,  numY: 2.280, numW: 1.354, numH: 1.212, num: '03', numColor: PURPLE, numAlpha: 70,
	  ringX: 6.886, ringY: 2.634, ringColor: 'FFFDEE', blobX: 6.796, blobY: 2.878, color: PURPLE,
	  year: '2026', yearX: 6.864, yearY: 4.218, yearW: 0.746, capX: 7.120, capY: 3.300, txtX: 7.224, txtY: 3.566 },
	{ numX: 11.142, numY: 3.703, numW: 1.452, numH: 1.212, num: '04', numColor: DARK, numAlpha: 80,
	  ringX: 9.495, ringY: 4.057, ringColor: 'DBDBDB', blobX: 9.405, blobY: 4.301, color: DARK,
	  year: '2027', yearX: 9.418, yearY: 5.641, yearW: 0.725, capX: 9.709, capY: 4.694, txtX: 9.813, txtY: 4.960 },
];

function slide11(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };
	heading(s, 'PLACEHOLDER',
		1.545, 0.503, 10.244, 1.178, 32, DARK, { align: 'center' });
	ring(s, 5.152, 1.102, 3.932, 0.531, PURPLE);

	BADGES.forEach((b) => {
		heading(s, b.num, b.numX, b.numY, b.numW, b.numH, 66, b.numColor,
			{ wrap: false, transparency: b.numAlpha });
		s.addShape('arc', {
			x: b.ringX - 0.179, y: b.ringY - 0.194, w: 2.585, h: 2.585,
			angleRange: [5.8, 3.4], fill: { type: 'none' },
			line: { color: GRAY, width: 2, dashType: 'dash' },
		});
		s.addShape('ellipse', {
			x: b.ringX, y: b.ringY, w: 2.198, h: 2.198,
			fill: { color: WHITE }, line: { color: b.ringColor, width: 10 },
		});
		poly(s, b.blobX, b.blobY, 2.043, 1.709, BADGE_BLOB, b.color);
		heading(s, b.year, b.yearX, b.yearY, b.yearW, 0.645, 18, WHITE, { align: 'center' });
		heading(s, 'Your Text', b.capX, b.capY, 1.762, 0.338, 14, WHITE, { align: 'center' });
		para(s, 'Lorem ipsum dolor sit amet', b.txtX, b.txtY, 1.583, 0.627,
			{ fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
	});
}

/* 12 - two jigsaw diamonds facing each other.  Each is four quarter diamonds; the cream
 * seams run from the centre out to the four corner notches and carry a round tab. */
const QUARTERS = [[0, -1], [1, 0], [0, 1], [-1, 0]];              // N, E, S, W
const SEAMS = [[-1, -1], [1, -1], [1, 1], [-1, 1]];               // NW, NE, SE, SW

function puzzleDiamond(slide, cx, cy, w, h, color) {
	QUARTERS.forEach(([dx, dy]) => {
		box(slide, 'diamond', cx + (dx * w) / 4 - w / 4, cy + (dy * h) / 4 - h / 4, w / 2, h / 2, color,
			{ shadow: shadow(BLACK, 0.05, 35, 10, 50) });
	});
	SEAMS.forEach(([sx, sy]) => {
		slide.addShape('line', {
			x: cx, y: cy, w: (sx * w) / 4, h: (sy * h) / 4, line: { color: CREAM, width: 3 },
		});
		const kx = cx + sx * 0.40 * (w / 4), ky = cy + sy * 0.40 * (h / 4);
		box(slide, 'ellipse', kx - 0.22, ky - 0.22, 0.44, 0.44, CREAM);
		box(slide, 'ellipse', kx - 0.18, ky - 0.18, 0.36, 0.36, color);
	});
}

const VALUE_CARDS = [
	{ x: 0.798,  y: 5.197, value: '225+', color: PURPLE, label: 'Value One' },
	{ x: 10.409, y: 1.137, value: '325+', color: DARK,   label: 'Value Two' },
];

function slide12(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };
	puzzleDiamond(s, 3.173, 3.753, 4.640, 4.880, PURPLE);
	puzzleDiamond(s, 10.150, 3.753, 4.646, 4.880, DARK);

	[[2.649, PURPLE, 'coins'], [9.619, DARK, 'code']].forEach(([x, color, glyph]) => {
		s.addShape('ellipse', { x, y: 3.224, w: 1.066, h: 1.066,
			fill: { color: WHITE }, line: { color: 'F2F2F2', width: 1 },
			shadow: shadow(BLACK, 0.1, 25, 25, 45) });
		icon(s, glyph, x + 0.316, 3.540, 0.434, color);
	});

	// exchange badge in the middle
	box(s, 'ellipse', 5.992, 3.075, 1.349, 1.349, WHITE);
	box(s, 'rightArrow', 6.533, 3.714, 0.563, 0.315, PURPLE);
	box(s, 'rightArrow', 6.238, 3.471, 0.563, 0.315, DARK, { rotate: 180 });

	VALUE_CARDS.forEach((c) => {
		box(s, 'roundRect', c.x, c.y, 2.126, 1.167, WHITE,
			{ rectRadius: 0.125, shadow: shadow(BLACK, 0.1, 25, 25, 45) });
		heading(s, c.value, c.x + 0.341, c.y + 0.12, 1.445, 0.64, 32, c.color, { align: 'center' });
		para(s, c.label, c.x + 0.257, c.y + 0.605, 1.611, 0.386,
			{ fontSize: 14, color: '404040', align: 'center' });
	});
}

/* 13 - stacked funnel */
const FUNNEL = [
	{ x: 9.098, y: 2.321, w: 1.249, color: DARK,   label: '1000' },
	{ x: 7.874, y: 3.299, w: 3.697, color: PURPLE, label: '3000' },
	{ x: 7.284, y: 4.277, w: 4.878, color: DARK,   label: '4000' },
	{ x: 6.667, y: 5.256, w: 6.112, color: PURPLE, label: '5000' },
];

function slide13(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };

	FUNNEL.forEach((b) => {
		s.addText(b.label, {
			shape: 'rect', x: b.x, y: b.y, w: b.w, h: 0.937,
			fill: { color: b.color }, line: NO_LINE,
			fontFace: HEAD, fontSize: 14, bold: true, color: WHITE,
			align: 'center', valign: 'middle', lineSpacingMultiple: 1.3,
		});
	});
	box(s, 'ellipse', 7.215, 2.621, 1.052, 1.052, DARK);
	heading(s, '+3000', 7.262, 2.957, 0.947, 0.38, 16, PURPLE, { bold: true, wrap: false });
	box(s, 'ellipse', 9.935, 1.737, 1.052, 1.052, PURPLE);
	heading(s, '+3000', 9.988, 2.073, 0.947, 0.38, 16, DARK, { bold: true, wrap: false });

	heading(s, 'Property Market Secures Financial Freedom', 0.779, 0.778, 5.216, 2.121, 40, DARK);
	ring(s, 3.271, 1.477, 2.793, 0.724, PURPLE);

	box(s, 'roundRect', 0.862, 3.433, 5.121, 1.742, DARK,
		{ rectRadius: 0.096, shadow: shadow(BLACK, 0.1, 80, 30, 35) });
	box(s, 'roundRect', 5.211, 3.718, 0.333, 0.333, WHITE, { rectRadius: 0.167 });
	box(s, 'upArrow', 5.302, 3.801, 0.151, 0.167, DARK);
	heading(s, '40,214', 1.113, 3.598, 2.288, 0.572, 28, CREAM, { bold: true });
	para(s, 'Pretium aenean pharetra magna ac placerat get et  gravida cum legit adispacing mountain legit cum',
		1.113, 4.303, 4.516, 0.602, { color: WHITE });
	para(s, 'lorem ipsum dolor sit amet, consectetuer adipiscing elit. aenean commodo ligula eget dolor. aenean massa. Cum Pretium aenean pharetra magna ac placerat get et ',
		0.762, 5.688, 5.075, 0.865);
}

/* 14 - two review cards */
const REVIEWS = [
	{ x: 0.450, cardColor: PURPLE, radius: 0.161, pct: '80%', pctW: 2.409, textX: 0.930,
	  starX: 0.995, dimLast: false, badgeX: 4.108, badgeColor: CREAM, badge: '100+',
	  badgeTextX: 4.196, badgeTextW: 0.724, badgeInk: BLACK },
	{ x: 6.771, cardColor: DARK,   radius: 0.196, pct: '75%', pctW: 2.246, textX: 7.223,
	  starX: 7.288, dimLast: true,  badgeX: 11.440, badgeColor: PURPLE, badge: '50+',
	  badgeTextX: 11.581, badgeTextW: 0.619, badgeInk: CREAM },
];

function slide14(pptx) {
	const s = pptx.addSlide();
	s.background = { color: CREAM };
	REVIEWS.forEach((r) => {
		box(s, 'roundRect', r.x, 0.56, 6.112, 6.38, r.cardColor,
			{ rectRadius: r.radius, shadow: shadow(PURPLE, 0.15, 25) });
		heading(s, r.pct, r.x + 0.483, 0.917, r.pctW, 1.313, 72, WHITE, { wrap: false });
		for (let i = 0; i < 5; i++) {
			box(s, 'star5', r.starX + i * 0.221, 3.75, 0.181, 0.181, WHITE,
				r.dimLast && i === 4 ? { fill: { color: WHITE, transparency: 50 } } : {});
		}
		para(s, 'lorem ipsum legit dolor sit amet, consectetuer et cum adipiscing elit. aenean commodo ligula eget dolor. aenean massa. cum sociis natoque pentibus cum adipiscing elit. ',
			r.textX, 4.164, 5.177, 0.865, { color: WHITE, align: 'justify' });
		box(s, 'ellipse', r.badgeX, 5.439, 0.9, 0.9, r.badgeColor);
		heading(s, r.badge, r.badgeTextX, 5.72, r.badgeTextW, 0.37, 16, r.badgeInk,
			{ align: 'center', wrap: false });
	});
}

/* 15 - closing slide */
const CONTACTS = [
	{ label: 'Website',  value: 'www.property.com', y: 3.735, labelX: 9.887, labelW: 1.152, valueX: 9.562, valueW: 1.803 },
	{ label: 'Email',    value: 'www.property.com', y: 4.612, labelX: 10.035, labelW: 0.856, valueX: 9.562, valueW: 1.803 },
	{ label: 'No.Phone', value: '+12 765 879 987',  y: 5.490, labelX: 9.794, labelW: 1.338, valueX: 9.716, valueW: 1.496 },
];

function slide15(pptx) {
	const s = pptx.addSlide();
	s.background = { color: PURPLE };
	box(s, 'ellipse', 8.811, 3.281, 3.306, 3.306, CREAM);
	heading(s, 'Thank You \u2014 Your Next Real Estate Opportunity',
		1.656, 1.35, 8.2, 3.13, 60, CREAM);
	sparkle(s, 0.853, 1.027, 1.466, 0.646, CREAM, 315);
	ring(s, 7.275, 3.712, 2.192, 0.464, CREAM, 1.5);
	heading(s, 'Contact Us', 7.718, 3.775, 1.306, 0.337, 14, CREAM, { align: 'center', wrap: false });

	CONTACTS.forEach((c) => {
		heading(s, c.label, c.labelX, c.y, c.labelW, 0.37, 16, BLACK, { align: 'center', wrap: false });
		para(s, c.value, c.valueX, c.y + 0.339, c.valueW, 0.303,
			{ color: BLACK, align: 'center', lineSpacingMultiple: 1, wrap: false });
	});
}

/* ------------------------------------------------------------------ build */

function build() {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
	pptx.layout = 'WIDE';
	pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };
	pptx.title = 'Understanding the Property Investment Landscape';

	[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08,
		slide09, slide10, slide11, slide12, slide13, slide14, slide15]
		.forEach((fn) => fn(pptx));

	const out = path.join(__dirname, '01bb4b9d-f971-45cc-b9b0-f98235925c84_grok_final.pptx');
	return pptx.writeFile({ fileName: out }).then(() => console.log('wrote', out));
}

build().catch((err) => { console.error(err); process.exit(1); });
