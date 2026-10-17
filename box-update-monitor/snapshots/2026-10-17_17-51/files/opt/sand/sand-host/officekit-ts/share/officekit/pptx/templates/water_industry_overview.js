/**
 * "Water Industry" deck — recreated with pptxgenjs.
 * Raster photos / 3D models in the original are replaced by flat placeholder
 * rectangles that keep the original position, size and corner radius.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette ---
const NAVY = '09254A'; // theme tx2  — headings
const SKY = '6DCFF6'; // theme accent2 — highlight word
const AQUA = 'A3E5F9'; // theme accent3
const MINT = '82EEE4'; // theme accent4
const GREY = '7A748A'; // body copy
const LIGHT = 'F5F5F5'; // card background
const WHITE = 'FFFFFF';
const RULE = 'D9D9D9'; // hairline rules (white lumMod 85%)
const RULE2 = 'BFBFBF'; // ring outlines (white lumMod 75%)
const GRAD = '7FE3EE'; // flat stand-in for the AQUA->MINT gradient
const SOFT = 'E8F5FA'; // flat stand-in for the pale blue card gradient
const FOOT = 'A8A8A8'; // footer text (tx1 @ 40% on white)

const TITLE_FONT = 'urbanist'; // theme major latin
const BODY_FONT = 'Nunito Light'; // theme minor latin

// ---------------------------------------------------------------- helpers ---
const rr = (adj, w, h) => adj * Math.min(w, h); // OOXML roundRect adj -> inches

/** Text box: PowerPoint anchors text boxes at the top, pptxgenjs centers them. */
function tx(sl, text, o) {
	sl.addText(text, Object.assign({ isTextBox: true, valign: 'top', fontFace: BODY_FONT }, o));
}

/** Grey 12pt paragraph copy (line spacing 135%, 9pt space before). */
function body(sl, x, y, w, h, text, o) {
	tx(sl, text, Object.assign({ x, y, w, h, fontSize: 12, color: GREY, lineSpacingMultiple: 1.35, paraSpaceBefore: 9 }, o));
}

/** Section heading, e.g. ['Climate', 1], [' Change and '], ['Water Scarcity ', 0, SKY]. */
function heading(sl, x, y, w, h, parts, o) {
	tx(
		sl,
		parts.map(([t, bold, color]) => ({ text: t, options: { bold: !!bold, color: color || NAVY } })),
		Object.assign({ x, y, w, h, fontFace: TITLE_FONT, fontSize: 48 }, o)
	);
}

/** Small "— Water Industry" kicker above every section heading. */
function kicker(sl, x, y) {
	sl.addShape('line', { x, y: y + 0.134, w: 0.206, h: 0, line: { color: SKY, width: 1.5 } });
	tx(sl, 'Water Industry', { x: x + 0.248, y, w: 1.365, h: 0.269, fontFace: TITLE_FONT, fontSize: 10, color: SKY });
}

/** Empty picture placeholder — the reference fills these with accent2 at 4% alpha. */
function photo(sl, x, y, w, h, radius) {
	sl.addShape('roundRect', { x, y, w, h, rectRadius: radius === undefined ? 0.28 : radius, fill: { color: SKY, transparency: 96 } });
}

/** Stand-in for one of the two embedded 3D models: a flat tinted block, captioned. */
function model(sl, x, y, w, h, shape, tone) {
	sl.addShape(shape, { x, y, w, h, rectRadius: shape === 'roundRect' ? 0.3 : undefined, fill: { color: tone } });
	tx(sl, '[image]', { x, y: y + h / 2 - 0.16, w, h: 0.32, align: 'center', fontSize: 11, color: WHITE });
}

/** Two overlapping rounded squares (one tilted 47deg) used as an icon badge. */
function iconBadge(sl, x, y, s) {
	const d = s * 0.0735;
	const ln = { color: SKY, width: 0.25 };
	sl.addShape('roundRect', { x: x + d, y, w: s, h: s, rectRadius: rr(0.35953, s, s), fill: { type: 'none' }, line: ln });
	sl.addShape('roundRect', {
		x: x + d, y: y - d, w: s, h: s * 1.1469, rotate: 47.14,
		rectRadius: rr(0.5, s, s), fill: { type: 'none' }, line: ln,
	});
	const g = s * 0.34;
	sl.addShape('roundRect', { x: x + d + (s - g) / 2, y: y + (s - g) / 2, w: g, h: g, rectRadius: g * 0.2, fill: { color: NAVY } });
}

/** Small diagonal "open link" arrow inside a circle. */
function arrowCircle(sl, x, y, d) {
	const a = d * 0.30;
	sl.addShape('ellipse', { x, y, w: d, h: d, fill: { color: WHITE } });
	sl.addShape('line', {
		x: x + (d - a) / 2, y: y + (d - a) / 2, w: a, h: a, flipV: true,
		line: { color: NAVY, width: 0.75, endArrowType: 'triangle' },
	});
}

/**
 * Chevron built from two strokes meeting at a point.
 * `a` is the half-span across the chevron, `b` the depth of the point.
 */
function chevron(sl, cx, cy, a, b, color, dir) {
	const ln = { color, width: 1 };
	if (dir === 'up' || dir === 'down') {
		const down = dir === 'down';
		sl.addShape('line', { x: cx - a, y: cy - b, w: a, h: b * 2, line: ln, flipV: !down });
		sl.addShape('line', { x: cx, y: cy - b, w: a, h: b * 2, line: ln, flipV: down });
	} else {
		const right = dir === 'right';
		sl.addShape('line', { x: cx - b, y: cy - a, w: b * 2, h: a, line: ln, flipV: !right });
		sl.addShape('line', { x: cx - b, y: cy, w: b * 2, h: a, line: ln, flipV: right });
	}
}

/** Small check mark (two strokes) used in the slide 12 bullet lists. */
function checkMark(sl, cx, cy, s, color) {
	sl.addShape('line', { x: cx - s * 0.45, y: cy, w: s * 0.35, h: s * 0.35, line: { color, width: 1.25 } });
	sl.addShape('line', { x: cx - s * 0.1, y: cy - s * 0.45, w: s * 0.55, h: s * 0.8, line: { color, width: 1.25 }, flipV: true });
}

/** Cubic bezier polyline: segs are [x1,y1,x2,y2,x,y] fractions of w/h. */
function curve(sl, x, y, w, h, start, segs, line) {
	const pts = [{ x: start[0] * w, y: start[1] * h, moveTo: true }];
	segs.forEach((s) =>
		pts.push({ x: s[4] * w, y: s[5] * h, curve: { type: 'cubic', x1: s[0] * w, y1: s[1] * h, x2: s[2] * w, y2: s[3] * h } })
	);
	sl.addShape('custGeom', { x, y, w, h, points: pts, fill: { type: 'none' }, line });
}

/** "Explore More" gradient pill. */
function pill(sl, x, y, w, h, label, size) {
	sl.addShape('roundRect', { x, y, w, h, rectRadius: rr(0.5, w, h), fill: { color: GRAD } });
	tx(sl, label, { x: x + 0.091, y: y + 0.053, w: w - 0.18, h: 0.286, align: 'center', fontFace: TITLE_FONT, fontSize: size || 11, color: WHITE });
}

/** Ring + progress arc + centred percentage (slides 4 and 15). */
function donut(sl, x, y, d, endAngle, label, labelSize) {
	sl.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: RULE, width: 4, transparency: 70 } });
	sl.addShape('arc', { x: x + 0.003, y: y + 0.003, w: d - 0.006, h: d - 0.006, angleRange: [270, endAngle], line: { color: SKY, width: 4 } });
	tx(sl, label, {
		x: x + d / 2 - 0.552, y: y + d / 2 - 0.219, w: 1.103, h: 0.438,
		align: 'center', fontFace: TITLE_FONT, fontSize: labelSize || 20, bold: true, color: NAVY,
	});
}

/** Footer pill (page number + site) that the slide master paints on slides 2-20. */
function footer(sl, num) {
	sl.addShape('roundRect', {
		x: 11.458, y: 6.966, w: 1.703, h: 0.334, rectRadius: rr(0.5, 1.703, 0.334),
		fill: { color: WHITE, transparency: 90 }, line: { color: RULE2, width: 0.5, transparency: 50 },
	});
	tx(sl, 'www.website.com   /', { x: 11.458, y: 7.006, w: 1.429, h: 0.252, align: 'right', fontSize: 9, color: FOOT });
	tx(sl, String(num), { x: 12.667, y: 7.006, w: 0.464, h: 0.252, align: 'right', fontSize: 9, color: FOOT });
}

// ------------------------------------------------------------ shared copy ---
const LOREM_LONG =
	'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ';
const LOREM_XL =
	'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felids, ultricies nec, ';

// ------------------------------------------------------------ slide 1..20 ---
function slide01(sl) {
	// Decorative rings, top-left and bottom-right.
	sl.addShape('ellipse', { x: -1.926, y: -3.982, w: 5.353, h: 5.353, fill: { type: 'none' }, line: { color: MINT, width: 1 } });
	sl.addShape('ellipse', { x: 8.809, y: 3.935, w: 5.353, h: 5.353, fill: { type: 'none' }, line: { color: MINT, width: 1 } });
	sl.addShape('ellipse', { x: 9.864, y: 4.016, w: 5.353, h: 5.353, fill: { color: MINT, transparency: 61 } });
	sl.addShape('ellipse', { x: 9.285, y: 4.267, w: 5.353, h: 5.353, fill: { color: AQUA, transparency: 48 } });
	// "Water Planet" 3D model.
	model(sl, 9.222, 4.086, 4.773, 4.806, 'ellipse', '15517A');

	heading(sl, 0.75, 2.347, 9.761, 1.407, [['Water', 0, SKY], [' Industry'], ['.', 0, SKY]], { fontSize: 96, lineSpacingMultiple: 0.8 });
	pill(sl, 10.999, 2.854, 1.584, 0.392, 'Explore More', 12);
	tx(sl, '/2025 PRESENT', { x: 0.75, y: 5.074, w: 2.053, h: 0.37, align: 'right', fontFace: TITLE_FONT, fontSize: 16, color: GREY, paraSpaceBefore: 9 });
	body(sl, 3.097, 5.074, 6.188, 0.717, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque', { fontSize: 14 });
}

function slide02(sl) {
	photo(sl, 0.747, 2.417, 11.84, 4.331, 0.33);
	sl.addShape('roundRect', { x: 0.75, y: 2.417, w: 11.833, h: 4.331, rectRadius: rr(0.07369, 11.833, 4.331), fill: { color: '000000', transparency: 55 } });

	kicker(sl, 0.886, 0.752);
	heading(sl, 0.749, 1.031, 5.585, 0.909, [['the '], ['Overview', 0, SKY]]);
	body(sl, 6.892, 1.13, 5.692, 0.629, LOREM_LONG);

	// Play button.
	sl.addShape('ellipse', { x: 6.335, y: 4.25, w: 0.664, h: 0.664, fill: { color: WHITE, transparency: 58 } });
	sl.addShape('triangle', { x: 6.577, y: 4.488, w: 0.218, h: 0.188, rotate: 90, fill: { color: WHITE } });

	// Caption pill.
	sl.addShape('roundRect', { x: 3.264, y: 6.042, w: 6.806, h: 0.427, rectRadius: rr(0.5, 6.806, 0.427), fill: { color: WHITE, transparency: 58 } });
	tx(sl, 'Plays a Crucial Role in Ensuring the Availability, Quality, and Sustainability', {
		x: 3.427, y: 6.087, w: 6.479, h: 0.337, align: 'center', fontFace: TITLE_FONT, fontSize: 14, color: WHITE,
	});
}

function slide03(sl) {
	photo(sl, 3.806, 4.083, 2.667, 2.667, 0.29);
	sl.addShape('roundRect', { x: 3.806, y: 4.083, w: 2.667, h: 2.667, rectRadius: rr(0.10729, 2.667, 2.667), fill: { color: '000000', transparency: 55 } });
	sl.addShape('roundRect', { x: 9.917, y: 4.083, w: 2.667, h: 2.667, rectRadius: rr(0.10729, 2.667, 2.667), fill: { color: LIGHT } });
	sl.addShape('roundRect', { x: 6.861, y: 4.083, w: 2.667, h: 2.667, rectRadius: rr(0.10729, 2.667, 2.667), fill: { color: LIGHT } });

	kicker(sl, 0.905, 0.774);
	heading(sl, 0.768, 1.053, 5.899, 1.717, [['Importance', 1], [' of Water in '], ['Society', 0, SKY], [' ']]);
	body(sl, 8.165, 1.453, 4.401, 0.901, LOREM_LONG + ' dis, ');
	tx(sl, 'Read More', { x: 8.165, y: 2.466, w: 1.339, h: 0.303, fontFace: TITLE_FONT, fontSize: 12, italic: true, color: NAVY });

	// Left column: title + prev/next buttons.
	tx(sl, 'Rising Temperatures ', { x: 0.953, y: 5.783, w: 1.655, h: 0.64, valign: 'bottom', fontFace: TITLE_FONT, fontSize: 16, bold: true, color: NAVY });
	sl.addShape('ellipse', { x: 2.933, y: 5.806, w: 0.277, h: 0.277, fill: { type: 'none' }, line: { color: RULE2, width: 0.75 } });
	chevron(sl, 3.0715, 5.9445, 0.047, 0.024, NAVY, 'left');
	sl.addShape('ellipse', { x: 2.933, y: 6.145, w: 0.277, h: 0.277, fill: { color: GRAD } });
	chevron(sl, 3.0715, 6.2835, 0.047, 0.024, WHITE, 'right');
	sl.addShape('ellipse', { x: 12.103, y: 6.145, w: 0.277, h: 0.277, fill: { color: GRAD } });
	chevron(sl, 12.2415, 6.2835, 0.047, 0.024, WHITE, 'right');

	// Card captions.
	body(sl, 4.009, 5.521, 2.26, 0.901, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo', { color: WHITE, valign: 'bottom' });
	body(sl, 7.065, 5.521, 2.26, 0.901, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo', { valign: 'bottom' });
	body(sl, 10.12, 5.354, 2.26, 0.629, LOREM_SHORT);

	// Middle card statistics.
	tx(sl, '12K', { x: 7.064, y: 4.278, w: 1.129, h: 0.707, fontFace: TITLE_FONT, fontSize: 36, color: SKY });
	tx(sl, 'Popular ', { x: 7.879, y: 4.554, w: 1.198, h: 0.337, valign: 'bottom', fontFace: TITLE_FONT, fontSize: 14, color: NAVY });
	sl.addShape('triangle', { x: 7.169, y: 5.366, w: 0.129, h: 0.096, fill: { color: SKY } });
	tx(sl, [
		{ text: '63% Reduced Product ', options: { color: RULE2 } },
		{ text: '638', options: { color: GREY } },
	], { x: 7.298, y: 5.296, w: 2.027, h: 0.236, fontSize: 8, paraSpaceBefore: 9 });

	photo(sl, 0.953, 4.278, 1.512, 1.009, 0.12);
	photo(sl, 10.12, 4.278, 1.51, 1.009, 0.12);
}

function slide04(sl) {
	photo(sl, 7.625, 1.003, 4.958, 5.494, 0.4);
	kicker(sl, 0.887, 1.326);
	heading(sl, 0.75, 1.605, 5.766, 2.524, [['Smart', 1], [' Water Management and '], ['IoT Integration ', 0, SKY]]);

	iconBadge(sl, 0.75, 4.737, 0.442);
	tx(sl, 'Water Distribution Systems', { x: 1.358, y: 4.79, w: 2.973, h: 0.337, fontFace: TITLE_FONT, fontSize: 14, color: NAVY });
	body(sl, 1.358, 5.273, 3.291, 0.901, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. ');

	sl.addShape('roundRect', { x: 6.605, y: 4.042, w: 1.764, h: 1.862, rectRadius: rr(0.18405, 1.764, 1.862), fill: { color: LIGHT } });
	tx(sl, 'Distribution', { x: 6.73, y: 4.204, w: 1.514, h: 0.303, align: 'center', fontFace: TITLE_FONT, fontSize: 12, color: NAVY });
	donut(sl, 6.945, 4.615, 1.084, 0, '56%', 16);
	tx(sl, 'Details', { x: 6.605, y: 6.0, w: 1.339, h: 0.303, fontFace: TITLE_FONT, fontSize: 12, underline: { style: 'sng' }, color: NAVY });

	pill(sl, 10.904, 5.837, 1.395, 0.392, 'Explore More');
}

function slide05(sl) {
	photo(sl, 0.75, 0.75, 2.849, 5.363, 0.3);
	photo(sl, 3.754, 1.387, 2.849, 5.363, 0.3);

	kicker(sl, 7.515, 1.621);
	heading(sl, 7.378, 1.9, 5.129, 1.717, [['Filtration', 1], [' and '], ['Sedimentation', 0, SKY], [' ']]);
	body(sl, 7.378, 3.899, 4.639, 0.901, LOREM_LONG + ' dis');

	[7.378, 8.111, 8.843].forEach((x) => iconBadge(sl, x, 5.39, 0.49));

	[
		[0.75, 6.267, 'Natural Particle Settling'],
		[3.754, 0.75, 'Removing Fine Impurities'],
	].forEach(([x, y, label]) => {
		sl.addShape('roundRect', { x, y, w: 2.849, h: 0.483, rectRadius: rr(0.5, 2.849, 0.483), fill: { color: LIGHT } });
		tx(sl, label, { x: x + 0.151, y: y + 0.073, w: 2.547, h: 0.337, align: 'center', fontFace: TITLE_FONT, fontSize: 14, color: NAVY });
	});
}

function slide06(sl) {
	sl.addShape('roundRect', { x: 0.752, y: 3.339, w: 5.698, h: 3.411, rectRadius: rr(0.09333, 5.698, 3.411), fill: { color: LIGHT } });
	sl.addShape('roundRect', { x: 6.783, y: 3.339, w: 2.734, h: 3.411, rectRadius: rr(0.12708, 2.734, 3.411), fill: { color: SOFT } });
	sl.addShape('roundRect', { x: 9.85, y: 3.339, w: 2.734, h: 3.411, rectRadius: rr(0.12708, 2.734, 3.411), fill: { color: LIGHT } });
	photo(sl, 4.141, 3.613, 2.069, 2.863, 0.22);

	kicker(sl, 0.886, 0.752);
	heading(sl, 0.749, 1.031, 6.282, 1.717, [['Climate', 1], [' Change and '], ['Water Scarcity ', 0, SKY]]);

	const cards = [
		{ x: 0.992, w: 2.849, title: 'Rising Temperatures ', copy: LOREM_MED, ch: 1.174, arrow: 2.077 },
		{ x: 7.022, w: 2.256, title: 'Melting Glaciers', copy: LOREM_SHORT, ch: 0.629, arrow: 8.89 },
		{ x: 10.089, w: 2.256, title: 'Increased Water ', copy: LOREM_SHORT, ch: 0.629, arrow: 11.957 },
	];
	cards.forEach((c) => {
		tx(sl, c.title, { x: c.x, y: 3.728, w: c.w, h: 0.37, fontFace: TITLE_FONT, fontSize: 16, bold: true, color: NAVY });
		body(sl, c.x, 4.312, c.w, c.ch, c.copy);
		tx(sl, 'Read More', { x: c.x, y: 6.015, w: 1.339, h: 0.303, fontFace: TITLE_FONT, fontSize: 12, underline: { style: 'sng' }, color: NAVY });
		arrowCircle(sl, c.arrow, 5.972, 0.388);
	});
}

function slide07(sl) {
	photo(sl, 0.75, 0.75, 11.833, 3.356, 0.25);
	sl.addShape('roundRect', { x: 0.75, y: 0.75, w: 11.833, h: 3.356, rectRadius: rr(0.07013, 11.833, 3.356), fill: { color: '000000', transparency: 55 } });

	kicker(sl, 1.376, 1.43);
	heading(sl, 1.239, 1.709, 8.453, 1.717, [['Wastewater ', 0, WHITE], ['Management', 1, WHITE], [' and ', 0, WHITE], ['Recycling', 0, SKY]]);
	arrowCircle(sl, 11.86, 3.374, 0.388);

	const bars = [
		{ x: 1.239, title: 'Sewage Treat Plants (STPs) ', pct: '93%', fill: 2.347, copy: LOREM_MED, ch: 0.901 },
		{ x: 5.041, title: 'Industrial Wastewater Treat', pct: '34%', fill: 1.076, copy: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo', ch: 0.629 },
		{ x: 8.844, title: 'Water Recycling and Reuse', pct: '73%', fill: 2.065, copy: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. ', ch: 0.901 },
	];
	bars.forEach((b) => {
		tx(sl, b.title, { x: b.x, y: 4.683, w: 3.251, h: 0.37, fontFace: TITLE_FONT, fontSize: 16, bold: true, color: NAVY });
		sl.addShape('line', { x: b.x + 0.101, y: 5.264, w: 2.68, h: 0, line: { color: 'F2F2F2', width: 3 } });
		sl.addShape('line', { x: b.x + 0.101, y: 5.264, w: b.fill, h: 0, line: { color: SKY, width: 3 } });
		tx(sl, b.pct, { x: b.x + 2.782, y: 5.129, w: 0.469, h: 0.269, align: 'right', fontFace: TITLE_FONT, fontSize: 10, color: NAVY });
		body(sl, b.x, 5.606, 3.251, b.ch, b.copy);
	});
}

function slide08(sl) {
	photo(sl, 9.2, 0.75, 3.383, 6.0, 0.35);
	kicker(sl, 0.887, 0.837);
	heading(sl, 0.75, 1.116, 6.545, 1.717, [['Economic', 1], [' Aspects of the Water '], ['Industry', 0, SKY]]);

	sl.addShape('line', { x: 0.75, y: 3.319, w: 7.546, h: 0, line: { color: RULE, width: 1 } });
	sl.addShape('line', { x: 4.523, y: 3.319, w: 0, h: 1.655, line: { color: RULE, width: 1 } });
	sl.addShape('line', { x: 0.75, y: 4.974, w: 7.546, h: 0, line: { color: RULE, width: 1 } });

	body(sl, 0.75, 3.56, 3.553, 1.174, LOREM_LONG + ' dis, ');
	body(sl, 4.813, 3.56, 3.483, 0.901, '\u201CLorem ipsum dolor sit amet, consectetu adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa.\u201D', {
		fontFace: 'Nunito ', bold: true, color: NAVY,
	});

	[
		[0.75, 1.453, 'Market Trends and Growth '],
		[4.813, 5.446, 'Public vs. Private Sector'],
	].forEach(([bx, tX, title]) => {
		iconBadge(sl, bx, 5.487, 0.49);
		tx(sl, title, { x: tX, y: 5.552, w: 2.85, h: 0.337, fontFace: TITLE_FONT, fontSize: 14, color: NAVY });
		body(sl, tX, 6.035, 2.85, 0.629, 'Lorem ipsum dolor sit amet, consecr adipiscing elit. Aenean commodo');
	});
}

function slide09(sl) {
	kicker(sl, 0.887, 0.762);
	heading(sl, 0.75, 1.041, 7.008, 0.909, [['Public', 1], [' vs. '], ['Private', 0, SKY]]);
	// "Astro Marine Backpack" 3D model.
	model(sl, 5.116, 2.399, 3.101, 4.453, 'roundRect', '4A5C86');

	const dot = (x, y) => sl.addShape('ellipse', { x, y, w: 0.286, h: 0.286, fill: { color: WHITE, transparency: 55 }, line: { color: SKY, width: 1 } });
	const elbow = (o) => sl.addShape('bentConnector3', Object.assign({ line: { color: SKY, width: 1, endArrowType: 'oval' }, fill: { type: 'none' } }, o));

	// Right hand callouts.
	[
		{ dot: [7.229, 3.156], elbow: { x: 7.515, y: 2.543, w: 1.155, h: 0.756, flipV: true }, tX: 8.915, tY: 2.358, bW: 3.244, bY: 2.824, copy: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula' },
		{ dot: [7.604, 5.512], elbow: { x: 7.89, y: 4.889, w: 1.56, h: 0.766, flipV: true }, tX: 9.665, tY: 4.714, bW: 2.918, bY: 5.18, copy: 'Lorem ipsum dolor sit amet, adipiscing elit. commodo ligula' },
	].forEach((c) => {
		dot(c.dot[0], c.dot[1]);
		elbow(c.elbow);
		tx(sl, 'Your Text Here One', { x: c.tX, y: c.tY, w: 2.849, h: 0.37, fontFace: TITLE_FONT, fontSize: 16, bold: true, color: NAVY });
		body(sl, c.tX, c.bY, c.bW, 0.629, c.copy);
	});

	// Left hand callouts (right aligned text).
	[
		{ dot: [6.261, 3.507], elbow: { x: 4.373, y: 3.064, w: 1.889, h: 0.587, rotate: 180 }, tX: 1.41, tY: 2.87, bX: 1.014, bW: 3.244, bY: 3.336, copy: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula' },
		{ dot: [6.786, 4.315], elbow: { x: 3.883, y: 4.458, w: 2.903, h: 0.508, rotate: 180, flipV: true }, tX: 0.871, tY: 4.759, bX: 1.135, bW: 2.585, bY: 5.225, copy: 'Lorem ipsum dolor sit amet, adipiscing elit. commodo ligula' },
	].forEach((c) => {
		dot(c.dot[0], c.dot[1]);
		elbow(c.elbow);
		tx(sl, 'Your Text Here One', { x: c.tX, y: c.tY, w: 2.849, h: 0.37, align: 'right', fontFace: TITLE_FONT, fontSize: 16, bold: true, color: NAVY });
		body(sl, c.bX, c.bY, c.bW, 0.629, c.copy, { align: 'right' });
	});
}

function slide10(sl) {
	kicker(sl, 0.887, 0.762);
	heading(sl, 0.75, 1.041, 7.008, 1.717, [['Environmental', 1], [' and Social '], ['Impact', 0, SKY]]);

	iconBadge(sl, 7.942, 3.348, 0.49);
	tx(sl, 'Sustainable Water Management Strategies ', { x: 7.942, y: 4.007, w: 4.642, h: 0.337, fontFace: TITLE_FONT, fontSize: 14, color: NAVY });
	body(sl, 7.942, 4.448, 4.642, 1.573, [
		{ text: LOREM_LONG.replace('penatibus et magnis', 'penatibus et magnis dis parturient montes, '), options: { breakLine: true } },
		{ text: 'Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, arcu. In enim justo, ' },
	]);
	pill(sl, 7.942, 6.358, 1.395, 0.392, 'Explore More');

	sl.addShape('roundRect', { x: 2.948, y: 5.075, w: 3.077, h: 1.675, rectRadius: rr(0.12708, 3.077, 1.675), fill: { color: LIGHT } });
	tx(sl, 'Pollution and Water Quality Issues ', { x: 3.169, y: 5.241, w: 2.552, h: 0.64, fontFace: TITLE_FONT, fontSize: 16, bold: true, color: NAVY });
	body(sl, 3.169, 5.955, 2.636, 0.629, 'Lorem ipsum dolor sit amet, consectetuer adipiscing. Aenean');

	[0.75, 2.605, 4.458].forEach((x) => photo(sl, x, 3.348, 1.855, 2.365, 0.18));
}

function slide11(sl) {
	// Faint background circles.
	[[-0.99, 4.368, 6.25], [8.152, 1.031, 7.642], [9.909, 2.788, 4.129]].forEach(([x, y, d]) =>
		sl.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: { color: RULE, width: 1 } })
	);

	// Phone mock-up.
	sl.addShape('roundRect', { x: 4.755, y: 2.55, w: 3.827, h: 7.818, rectRadius: 0.6, fill: { color: WHITE } });
	sl.addShape('roundRect', { x: 4.992, y: 2.788, w: 3.353, h: 7.344, rectRadius: 0.5, fill: { color: 'F2F4F3' } });
	[[8.544, 4.356, 0.106, 0.906], [4.687, 4.143, 0.144, 0.596], [4.683, 4.852, 0.144, 0.596], [4.698, 3.572, 0.114, 0.341]].forEach(([x, y, w, h]) =>
		sl.addShape('roundRect', { x, y, w, h, rectRadius: w / 2, fill: { color: RULE } })
	);
	photo(sl, 4.915, 2.683, 3.507, 7.552, 0.45);

	kicker(sl, 0.887, 0.762);
	heading(sl, 0.75, 1.041, 7.008, 0.909, [['Water '], ['Infrastructure', 0, SKY]]);

	[
		[10.651, 1.92, 'Ashley E. Corby'],
		[10.651, 4.076, 'Lisa P. Hill'],
		[2.464, 5.317, 'Esther R. Lopez'],
	].forEach(([x, y, name]) => {
		tx(sl, name, { x, y, w: 1.933, h: 0.337, fontFace: TITLE_FONT, fontSize: 14, color: NAVY });
		tx(sl, 'Your Jobdesc Here', { x: x - 0.01, y: y + 0.387, w: 1.933, h: 0.303, fontSize: 12, color: GREY, paraSpaceBefore: 9 });
	});

	[[9.682, 1.093], [9.537, 3.914], [2.455, 4.097]].forEach(([x, y]) => photo(sl, x, y, 1.014, 1.014, 0.5));
}

function slide12(sl) {
	kicker(sl, 0.887, 0.762);
	heading(sl, 0.75, 1.041, 7.008, 1.717, [['Resilience', 1], [' and climate change '], ['adaptation', 0, SKY]]);
	body(sl, 8.933, 1.667, 3.65, 0.901, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoq');

	[
		{ x: 0.75, tab: '01', title: 'Resilience ' },
		{ x: 4.792, tab: '02', title: 'Adaptation ' },
	].forEach((c) => {
		sl.addShape('round2SameRect', { x: c.x + 2.658, y: 3.401, w: 0.665, h: 0.328, fill: { color: GRAD } });
		tx(sl, c.tab, { x: c.x + 2.658, y: 3.435, w: 0.665, h: 0.26, align: 'center', fontFace: TITLE_FONT, fontSize: 12, bold: true, color: WHITE });
		sl.addShape('roundRect', { x: c.x, y: 3.73, w: 3.75, h: 3.02, rectRadius: rr(0.1262, 3.75, 3.02), fill: { color: LIGHT } });
		iconBadge(sl, c.x + 0.393, 4.081, 0.49);
		tx(sl, c.title, { x: c.x + 0.393, y: 4.74, w: 2.93, h: 0.404, fontFace: TITLE_FONT, fontSize: 18, color: NAVY });
		[
			[5.56, 'Lorem ipsum dolor sit amet'],
			[6.06, 'Aenean commodo ligula eget'],
		].forEach(([y, label]) => {
			sl.addShape('ellipse', { x: c.x + 0.393, y, w: 0.339, h: 0.338, fill: { color: AQUA, transparency: 85 } });
			checkMark(sl, c.x + 0.5625, y + 0.169, 0.16, 'C89B7B');
			tx(sl, label, { x: c.x + 0.789, y: y + 0.018, w: 2.568, h: 0.303, fontSize: 12, color: GREY, lineSpacingMultiple: 1.0 });
		});
	});

	photo(sl, 8.833, 3.73, 3.75, 3.02, 0.38);
}

function slide13(sl) {
	kicker(sl, 0.887, 1.193);
	heading(sl, 0.75, 1.472, 4.894, 1.717, [['Water', 1], [' Industry '], ['Trends', 0, SKY]]);
	body(sl, 7.256, 1.193, 5.328, 1.174, LOREM_XL.replace('ultricies nec, ', 'ultricies nec, pellentesque eu.'));

	[3.626, 5.05, 5.764].forEach((y) => sl.addShape('line', { x: 7.256, y, w: 5.33, h: 0, line: { color: RULE, width: 1 } }));

	const rows = [
		{ y: 3.083, label: 'Emerging technologies and solutions', open: false },
		{ y: 3.798, label: 'Data analytics and digitalization', open: true },
		{ y: 5.222, label: 'Circular economy and resource recovery', open: false },
		{ y: 5.937, label: 'Sustainability and environmental stewardship', open: false },
	];
	rows.forEach((r) => {
		tx(sl, r.label, { x: 7.256, y: r.y, w: 4.777, h: 0.37, fontFace: TITLE_FONT, fontSize: 16, color: NAVY });
		const cy = r.y + 0.175;
		if (r.open) {
			sl.addShape('ellipse', { x: 12.306, y: r.y + 0.047, w: 0.277, h: 0.277, fill: { color: GRAD } });
			chevron(sl, 12.4445, cy + 0.047, 0.047, 0.024, WHITE, 'up');
		} else {
			sl.addShape('ellipse', { x: 12.31, y: r.y + 0.037, w: 0.277, h: 0.277, fill: { type: 'none' }, line: { color: RULE2, width: 0.75 } });
			chevron(sl, 12.4485, cy + 0.037, 0.047, 0.024, NAVY, 'down');
		}
	});
	body(sl, 7.256, 4.249, 5.328, 0.629, LOREM_MED);

	photo(sl, 0.75, 3.626, 5.323, 2.681, 0.3);
}

// Sparkline drawn on the slide-14 stat card (cubic segments, fractions of w/h).
const SPARKLINE = [
	[0.021, 0.531, 0.048, 0.622, 0.079, 0.659], [0.1, 0.685, 0.124, 0.682, 0.146, 0.647],
	[0.165, 0.616, 0.183, 0.561, 0.199, 0.483], [0.226, 0.341, 0.248, 0.194, 0.279, 0.1],
	[0.289, 0.068, 0.301, 0.043, 0.314, 0.026], [0.354, -0.03, 0.4, 0.001, 0.43, 0.156],
	[0.448, 0.253, 0.461, 0.382, 0.474, 0.51], [0.48, 0.572, 0.487, 0.63, 0.494, 0.68],
	[0.519, 0.852, 0.552, 0.922, 0.594, 0.888], [0.619, 0.868, 0.643, 0.784, 0.668, 0.723],
	[0.699, 0.646, 0.734, 0.63, 0.766, 0.689], [0.792, 0.738, 0.806, 0.915, 0.833, 0.955],
	[0.848, 0.977, 0.863, 0.954, 0.877, 0.937], [0.91, 0.896, 0.927, 1.024, 0.957, 0.996],
	[0.965, 0.989, 0.972, 0.966, 0.979, 0.951], [0.987, 0.936, 1.0, 0.936, 1.0, 0.973],
];

function slide14(sl) {
	sl.addShape('roundRect', { x: 0.75, y: 3.451, w: 3.734, h: 3.299, rectRadius: rr(0.12552, 3.734, 3.299), fill: { color: LIGHT } });
	sl.addShape('roundRect', { x: 8.85, y: 3.451, w: 3.734, h: 3.299, rectRadius: rr(0.12552, 3.734, 3.299), fill: { color: LIGHT } });
	photo(sl, 4.8, 3.451, 3.734, 3.299, 0.41);

	kicker(sl, 0.887, 0.762);
	heading(sl, 0.75, 1.041, 6.035, 1.717, [['Water '], ['education', 1], [' and '], ['awareness', 0, SKY]]);
	body(sl, 7.256, 1.173, 5.328, 1.174, LOREM_XL);

	// Left card: copy, big number, sparkline and callout bubble.
	body(sl, 1.074, 3.781, 3.085, 0.629, 'Lorem ipsum dolor sit amet, adipiscing elit. Aenean commodo ligula eget');
	tx(sl, '64%', { x: 1.074, y: 4.507, w: 3.085, h: 0.438, fontFace: TITLE_FONT, fontSize: 20, bold: true, color: NAVY });
	curve(sl, 1.135, 5.557, 3.025, 0.863, [0, 0.396], SPARKLINE, { color: SKY, width: 0.75 });
	sl.addShape('wedgeRoundRectCallout', { x: 3.048, y: 5.211, w: 1.002, h: 0.455, fill: { color: GRAD } });
	tx(sl, '$632', { x: 3.201, y: 5.217, w: 0.695, h: 0.303, align: 'center', fontFace: TITLE_FONT, fontSize: 12, bold: true, color: WHITE });
	tx(sl, 'Lorem ipsum', { x: 3.16, y: 5.44, w: 0.777, h: 0.219, align: 'center', fontSize: 7, color: WHITE, paraSpaceBefore: 9 });

	// Right card.
	tx(sl, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean ', {
		x: 9.174, y: 3.781, w: 3.085, h: 1.178, fontFace: TITLE_FONT, fontSize: 16, color: NAVY,
	});
	body(sl, 9.174, 5.791, 3.085, 0.629, 'Lorem ipsum dolor sit amet, adipiscing elit. Aenean commodo ligula eget');
}

function slide15(sl) {
	kicker(sl, 0.887, 0.762);
	heading(sl, 0.75, 1.041, 6.035, 0.909, [['Annual', 1], [' Objectives']]);

	[
		{ x: 0.75, cx: 0.649, end: 96.7, label: '56%' },
		{ x: 3.282, cx: 3.18, end: 24.9, label: '39%' },
	].forEach((d) => {
		donut(sl, d.x, 2.501, 1.649, d.end, d.label);
		tx(sl, 'Your Text Here #1', { x: d.cx, y: 4.268, w: 1.851, h: 0.303, align: 'center', fontFace: TITLE_FONT, fontSize: 12, color: NAVY });
	});
	body(sl, 0.75, 5.465, 5.623, 1.174, LOREM_XL);

	// Venn-style trio of gradient bubbles around a hairline circle.
	sl.addShape('ellipse', { x: 8.359, y: 3.288, w: 3.222, h: 3.222, fill: { type: 'none' }, line: { color: RULE, width: 1 } });
	[[8.995, 1.923], [7.325, 4.815], [10.665, 4.812]].forEach(([x, y]) => {
		sl.addShape('ellipse', { x, y, w: 1.95, h: 1.95, fill: { color: GRAD } });
		sl.addShape('roundRect', { x: x + 0.82, y: y + 0.426, w: 0.31, h: 0.31, rectRadius: 0.06, fill: { color: WHITE, transparency: 25 } });
		tx(sl, 'Water Resources', { x: x + 0.24, y: y + 0.885, w: 1.47, h: 0.64, align: 'center', fontFace: TITLE_FONT, fontSize: 16, color: WHITE });
	});
	tx(sl, '39%', { x: 9.497, y: 4.451, w: 1.103, h: 0.572, align: 'center', fontFace: TITLE_FONT, fontSize: 28, color: NAVY });
	tx(sl, 'Your Text Here', { x: 9.122, y: 5.043, w: 1.851, h: 0.303, align: 'center', fontFace: TITLE_FONT, fontSize: 12, color: NAVY });
}

// The large "growth" curve that sweeps across slide 16.
const GROWTH = [
	[0.121, 0.759, 0.132, 0.605, 0.182, 0.498], [0.233, 0.392, 0.267, 0.373, 0.302, 0.36],
	[0.31, 0.357, 0.318, 0.356, 0.326, 0.358], [0.342, 0.362, 0.356, 0.371, 0.374, 0.393],
	[0.379, 0.399, 0.384, 0.406, 0.389, 0.415], [0.423, 0.472, 0.462, 0.561, 0.496, 0.642],
	[0.505, 0.664, 0.514, 0.682, 0.523, 0.696], [0.547, 0.736, 0.571, 0.747, 0.602, 0.689],
	[0.625, 0.645, 0.645, 0.575, 0.664, 0.497], [0.673, 0.462, 0.681, 0.425, 0.69, 0.388],
	[0.716, 0.272, 0.743, 0.15, 0.774, 0.078], [0.782, 0.059, 0.791, 0.045, 0.799, 0.032],
	[0.813, 0.011, 0.827, -0.002, 0.842, 0.0], [0.85, 0.001, 0.859, 0.012, 0.867, 0.028],
	[0.882, 0.058, 0.901, 0.133, 0.919, 0.192], [0.948, 0.285, 0.981, 0.362, 1.0, 0.227],
];

function slide16(sl) {
	curve(sl, 6.329, 1.685, 7.004, 4.859, [0, 1], GROWTH, { color: SKY, width: 0.75 });
	sl.addShape('line', { x: 0, y: 6.544, w: 13.333, h: 0, line: { color: RULE, width: 1 } });

	[
		{ lx: 5.614, ly: 5.578, dx: 6.597, dy: 5.882, v: '13K' },
		{ lx: 7.484, ly: 3.061, dx: 8.468, dy: 3.365, v: '26K' },
		{ lx: 9.479, ly: 4.692, dx: 10.171, dy: 5.133, v: '18K' },
		{ lx: 11.151, ly: 1.308, dx: 12.134, dy: 1.612, v: '58K' },
	].forEach((m) => {
		sl.addShape('ellipse', { x: m.dx, y: m.dy, w: 0.2, h: 0.2, fill: { color: GRAD } });
		tx(sl, m.v, { x: m.lx, y: m.ly, w: 0.983, h: 0.404, align: 'right', fontFace: TITLE_FONT, fontSize: 18, color: NAVY });
	});

	kicker(sl, 0.887, 2.419);
	heading(sl, 0.75, 2.698, 6.035, 0.909, [['Growth ', 1], ['Industry']]);
	body(sl, 0.75, 3.907, 5.579, 1.174, LOREM_XL.replace('ultricies nec, ', 'ultricies nec, pellentesque eu, pretium quis, sem. '));
}

function slide17(sl) {
	kicker(sl, 0.887, 0.782);
	heading(sl, 0.75, 1.061, 6.415, 0.909, [['Percentage ', 1], ['Industry']]);
	body(sl, 7.78, 1.262, 4.803, 0.629, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean ligula eget dolor. Aenean massa. Cum sociis natoque ');

	[
		{ x: 0.75, y: 2.761, icon: [3.417, 3.289], img: [0.997, 2.978] },
		{ x: 4.986, y: 4.938, icon: [7.653, 5.466], img: [5.233, 5.155] },
	].forEach((c) => {
		sl.addShape('roundRect', { x: c.x, y: c.y, w: 7.597, h: 1.812, rectRadius: rr(0.16667, 7.597, 1.812), fill: { color: LIGHT } });
		photo(sl, c.img[0], c.img[1], 2.181, 1.378, 0.14);
		iconBadge(sl, c.icon[0], c.icon[1], 0.673);
		[
			['55% ', 'of businesses globally offer some capacity for remote work ', 0],
			['18%', 'of the workforce telecommute on a full-time basis ', 0.685],
		].forEach(([pct, label, dy]) => {
			tx(sl, pct, { x: c.x + 3.493, y: c.y + 0.391 + dy, w: 0.768, h: 0.37, align: 'right', fontFace: TITLE_FONT, fontSize: 16, color: NAVY });
			tx(sl, label, { x: c.x + 4.375, y: c.y + 0.261 + dy, w: 2.76, h: 0.629, fontSize: 12, color: GREY });
		});
	});
}

function slide18(sl) {
	sl.addShape('roundRect', { x: 0.75, y: 3.506, w: 9.361, h: 2.937, rectRadius: rr(0.13219, 9.361, 2.937), fill: { color: LIGHT } });
	photo(sl, 8.674, 0.762, 3.909, 5.988, 0.4);

	kicker(sl, 0.887, 1.057);
	heading(sl, 0.75, 1.336, 6.035, 1.717, [['Talks', 1], [' for Conference']]);

	// Five "user" pictograms — the first three are highlighted.
	[1.322, 2.553, 3.784, 5.016, 6.247].forEach((x, i) => {
		const color = i < 3 ? GRAD : RULE;
		sl.addShape('ellipse', { x: x + 0.109, y: 3.871, w: 0.656, h: 0.656, fill: { color } });
		sl.addShape('pie', { x, y: 4.582, w: 0.874, h: 0.874, angleRange: [180, 360], fill: { color } });
	});

	tx(sl, '65% - 80%', { x: 7.769, y: 4.026, w: 1.971, h: 0.572, fontFace: TITLE_FONT, fontSize: 28, color: NAVY });
	tx(sl, 'User Active', { x: 7.769, y: 4.549, w: 1.727, h: 0.37, fontFace: TITLE_FONT, fontSize: 16, color: NAVY });
	body(sl, 1.121, 5.449, 6.096, 0.629, LOREM_LONG + ' dis');
}

function slide19(sl) {
	sl.addShape('rect', { x: 0, y: 3.077, w: 6.079, h: 4.423, fill: { color: NAVY } });
	sl.addShape('roundRect', { x: 4.845, y: 2.327, w: 7.738, h: 4.423, rectRadius: rr(0.08766, 7.738, 4.423), fill: { color: LIGHT } });

	tx(sl, '75%', { x: 1.417, y: 3.587, w: 1.889, h: 0.874, fontFace: TITLE_FONT, fontSize: 44, bold: true, color: WHITE, lineSpacingMultiple: 1.1, paraSpaceBefore: 6 });
	sl.addShape('line', { x: 1.528, y: 4.608, w: 2.226, h: 0, line: { color: WHITE, width: 1, transparency: 50 } });
	tx(sl, 'A wonderful serenity has taken possession of my entire', { x: 1.417, y: 5.336, w: 2.337, h: 0.995, fontSize: 14, color: WHITE });

	sl.addChart(
		'bar',
		[{
			name: 'Series 1',
			labels: ['Unplugging after work', 'Loneliness', 'Collaborating', 'Taking vacation time', 'Staying motivated'],
			values: [22, 19, 17, 8, 10],
		}],
		{
			x: 5.259, y: 2.667, w: 6.91, h: 3.742,
			barDir: 'col', barGapWidthPct: 250, barOverlapPct: -27,
			chartColors: [SKY], showLegend: false, showTitle: false,
			valAxisMaxVal: 25, valAxisMajorUnit: 5, valAxisLineShow: false,
			valAxisLabelColor: '595959', valAxisLabelFontFace: BODY_FONT, valAxisLabelFontSize: 10,
			valGridLine: { color: 'E8E8E8', size: 0.75 },
			catAxisLineShow: false, catAxisLabelColor: '404040', catAxisLabelFontFace: BODY_FONT, catAxisLabelFontSize: 10,
			plotArea: { fill: { color: WHITE, transparency: 100 } },
			chartArea: { fill: { color: WHITE, transparency: 100 } },
		}
	);

	sl.addShape('wedgeRoundRectCallout', { x: 10.216, y: 2.667, w: 1.766, h: 1.212, fill: { color: GRAD } });
	tx(sl, '+100', { x: 10.33, y: 2.763, w: 1.538, h: 0.804, align: 'center', fontFace: TITLE_FONT, fontSize: 36, color: WHITE, lineSpacingMultiple: 1.1, paraSpaceBefore: 6 });
	tx(sl, 'Value Title', { x: 10.334, y: 3.323, w: 1.534, h: 0.431, align: 'center', fontSize: 14, color: WHITE });

	kicker(sl, 0.887, 0.762);
	heading(sl, 0.75, 1.041, 6.035, 0.909, [['Chart ', 1], ['Infographic']]);
}

function slide20(sl) {
	heading(sl, 0.75, 1.041, 11.833, 1.447, [['Thank ', 1], ['You']], { fontSize: 80 });
	photo(sl, 0.75, 2.893, 6.506, 3.483, 0.3);

	tx(sl, 'Get in Touch', { x: 8.137, y: 3.224, w: 4.267, h: 0.404, fontFace: TITLE_FONT, fontSize: 18, color: SKY });
	[
		[4.102, 'BusinessDiscussions.com'],
		[4.829, '618-623-5816'],
		[5.555, '1810 College View Stlouis, IL 63101'],
	].forEach(([y, label]) => {
		iconBadge(sl, 8.137, y, 0.49);
		tx(sl, label, { x: 8.877, y: y + 0.077, w: 3.401, h: 0.337, fontSize: 14, color: GREY, lineSpacingMultiple: 1.0 });
	});
}

// -------------------------------------------------------------------- run ---
const BUILDERS = [
	slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
	slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20,
];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE_13x7_5', width: 13.3333, height: 7.5 });
pptx.layout = 'WIDE_13x7_5';
pptx.author = 'pptxgenjs';
pptx.title = 'Water Industry';
pptx.theme = { headFontFace: TITLE_FONT, bodyFontFace: BODY_FONT };

BUILDERS.forEach((build, i) => {
	const sl = pptx.addSlide();
	sl.background = { color: WHITE };
	build(sl);
	if (i > 0) footer(sl, i + 1); // the cover slide hides the master footer
});

pptx.writeFile({ fileName: path.join(__dirname, '0905a089-c8bc-45e6-b939-8060b7ddab8f_grok_final.pptx') })
	.then((f) => console.log('wrote', f));
