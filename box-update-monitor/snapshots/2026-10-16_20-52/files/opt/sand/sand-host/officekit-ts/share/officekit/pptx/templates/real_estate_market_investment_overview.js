/**
 * "Horizon Estates - Real Estate Presentation" rebuilt with pptxgenjs.
 * 20 slides, 13.333 x 7.5 in (16:9).
 *
 * Photographs in the source deck are recreated as flat grey placeholder
 * rectangles; vector icons are recreated as outlined placeholder squares.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ */
/* palette + typography                                                */
/* ------------------------------------------------------------------ */
const LIGHT = 'F7F7F7';       // light slide background
const DARK = '2C2624';        // dark slide background / dark cards
const ORANGE = 'FE8049';      // accent
const DEEP_ORANGE = 'F44B01'; // icon accent used on orange fills
const WHITE = 'FFFFFF';
const OFFWHITE = 'F2F2F2';
const BLACK = '000000';
const GREY = '808080';        // body copy on light slides
const ICON_GREY = 'DEDEDE';   // inactive icon outline
const RULE = 'BFBFBF';
const PH_FILL = 'CFCFCF';     // image placeholder fill
const PH_TEXT = '9C9C9C';     // image placeholder caption

const TITLE = 'Outfit Light';
const BODY = 'Poppins';
const BODY_M = 'Poppins Medium';
const BODY_B = 'Poppins Bold';

const DROP = { type: 'outer', color: BLACK, blur: 20, offset: 10, angle: 90, opacity: 0.1 };
const SOFT = { type: 'outer', color: BLACK, blur: 15, offset: 3, angle: 45, opacity: 0.2 };
const GLOW = { type: 'outer', color: ORANGE, blur: 20, offset: 5, angle: 90, opacity: 0.2 };

// pptxgenjs text margins are [left, right, bottom, top] in points
const TIGHT = [3.6, 3.6, 1.8, 1.8];

/* ------------------------------------------------------------------ */
/* tiny helpers                                                        */
/* ------------------------------------------------------------------ */

/** text box: source deck uses top-anchored, auto-sized boxes everywhere */
function txt(s, str, x, y, w, h, o) {
	s.addText(str, Object.assign({
		x: x, y: y, w: w, h: h,
		fontFace: BODY, fontSize: 11, color: GREY, valign: 'top',
	}, o));
}

function shape(s, kind, x, y, w, h, o) {
	s.addShape(kind, Object.assign({ x: x, y: y, w: w, h: h }, o));
}

/** flat rectangle standing in for a photograph */
function photo(s, x, y, w, h) {
	shape(s, 'rect', x, y, w, h, { fill: { color: PH_FILL } });
	if (w >= 1.2 && h >= 0.5) {
		txt(s, '[image]', x, y + h / 2 - 0.18, w, 0.36,
			{ fontSize: 12, color: PH_TEXT, align: 'center' });
	}
}

/** outlined square standing in for a vector icon */
function icon(s, x, y, w, h, color) {
	shape(s, 'roundRect', x, y, w, h, {
		rectRadius: Math.min(w, h) * 0.22,
		line: { color: color, width: 1.5 },
	});
}

/**
 * Elbow/straight connector drawn as an open polyline through absolute points.
 * `opts` may carry {color, dash, endCap, startCap} where the caps are
 * 'triangle' (arrow head) or 'oval' (dot).
 */
function polyline(s, pts, opts) {
	const xs = pts.map(p => p[0]);
	const ys = pts.map(p => p[1]);
	const x = Math.min.apply(null, xs);
	const y = Math.min.apply(null, ys);
	shape(s, 'custGeom', x, y, Math.max.apply(null, xs) - x || 0.01, Math.max.apply(null, ys) - y || 0.01, {
		points: pts.map(p => ({ x: p[0] - x, y: p[1] - y })),
		line: {
			color: opts.color, width: 1, dashType: opts.dash || 'solid',
			beginArrowType: opts.startCap, endArrowType: opts.endCap,
		},
	});
}

/** small 45-degree "go" arrow used inside the call-to-action buttons */
function arrow(s, x, y, size, color) {
	shape(s, 'line', x, y, size, size, {
		flipV: true, line: { color: color, width: 1, endArrowType: 'triangle' },
	});
}

/** header + footer that repeat on (almost) every slide */
function chrome(s, dark, withFooter) {
	const head = dark ? WHITE : BLACK;
	const sub = dark ? OFFWHITE : GREY;
	txt(s, 'Horizon Estates', 0.552, 0.281, 1.373, 0.303, { fontFace: TITLE, fontSize: 12, color: head });
	txt(s, 'Real Estate Presentation', 10.677, 0.281, 2.107, 0.303,
		{ fontFace: TITLE, fontSize: 12, color: sub, align: 'right' });
	if (withFooter !== false) {
		txt(s, 'www.yourgreatsite.com', 0.552, 6.911, 2.573, 0.303,
			{ fontFace: TITLE, fontSize: 12, color: sub });
	}
}

/** centred "Infographic Section" heading shared by slides 9-18 */
function sectionTitle(s, dark) {
	txt(s, 'Infographic Section', 3.938, 0.814, 5.457, 0.774,
		{ fontFace: TITLE, fontSize: 40, color: dark ? WHITE : DARK, align: 'center' });
}

/** solid button: pill of colour + label (+ optional arrow) */
function button(s, label, x, y, w, h, fill, tx, ty, tw) {
	shape(s, 'rect', x, y, w, h, { fill: { color: fill }, shadow: GLOW });
	txt(s, label, tx, ty, tw, 0.303, { fontFace: BODY_M, fontSize: 12, color: WHITE });
	arrow(s, x + w - 0.403, y + 0.091, 0.26, WHITE);
}

/* ------------------------------------------------------------------ */
/* repeated copy                                                       */
/* ------------------------------------------------------------------ */
const LOREM_SHORT = 'Lorem ipsum dolor sit, consectetur amet sit adipisicing elit, sed do';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ';
const LOREM_SECTOR = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Mae porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus.';

/* ------------------------------------------------------------------ */
/* slides                                                              */
/* ------------------------------------------------------------------ */

/* 1 — cover */
function slide01(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	photo(s, 7.333, 1.0, 5.333, 5.5);

	// three statistic chips across the top of the photo
	[
		{ x: 6.530, w: 1.472, value: '50k+', label: 'Our Partner', lw: 1.147 },
		{ x: 8.214, w: 1.569, value: '250k+', label: 'Happy Client', lw: 1.506 },
		{ x: 9.979, w: 2.131, value: '$500B+', label: 'Annual Transactions', lw: 1.877 },
	].forEach(c => {
		shape(s, 'rect', c.x, 1.330, c.w, 0.774, { fill: { color: DARK }, shadow: DROP });
		txt(s, c.value, c.x + 0.195, 1.355, 1.082, 0.467,
			{ fontFace: BODY_M, fontSize: 16, color: WHITE, lineSpacingMultiple: 1.5 });
		txt(s, c.label, c.x + 0.195, 1.643, c.lw, 0.352,
			{ fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
	});

	txt(s, 'Real Estate Market', 0.683, 1.180, 3.423, 1.582, { fontFace: TITLE, fontSize: 44, color: BLACK });
	txt(s, 'Lorem ipsum dolor sit, consectetur elit, sed do  adipisicing eiusmod incid. Maecenas porttitor congue. ',
		0.683, 2.846, 3.171, 0.908, { lineSpacingMultiple: 1.5 });
	button(s, 'Get in Touch', 0.752, 4.118, 1.731, 0.443, ORANGE, 0.868, 4.188, 1.247);

	// three thumbnails + agent card
	[0.683, 2.046, 3.409].forEach(x => photo(s, x, 5.263, 1.242, 1.242));
	shape(s, 'rect', 6.530, 5.484, 3.231, 0.801, { fill: { color: WHITE }, shadow: DROP });
	txt(s, 'Gerald Brown', 6.725, 5.545, 1.482, 0.375,
		{ fontFace: BODY_M, fontSize: 12, color: BLACK, lineSpacingMultiple: 1.5 });
	txt(s, 'Property Manager', 6.725, 5.802, 1.640, 0.352, { fontSize: 10.5, lineSpacingMultiple: 1.5 });
	photo(s, 9.038, 5.585, 0.588, 0.588);
}

/* 2 — introduction */
function slide02(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	photo(s, 0.667, 1.016, 5.533, 2.532);
	photo(s, 0.667, 3.980, 5.533, 2.532);

	txt(s, 'Introduction to Real Estate Market', 7.280, 1.739, 5.046, 1.447,
		{ fontFace: TITLE, fontSize: 40, color: BLACK });
	txt(s, 'Lorem ipsum dolor sit, consectetur elit, sed do  adipisicing eiusmod incid. Maecenas porttitor congue massa. ',
		7.280, 3.301, 5.046, 0.630, { lineSpacingMultiple: 1.5 });
	txt(s, 'Exploring the World of Real Estate', 7.280, 4.046, 5.046, 0.352,
		{ fontFace: BODY_M, color: DARK, lineSpacingMultiple: 1.5 });
	txt(s, 'Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero. Lorem',
		7.280, 4.367, 5.046, 0.630, { lineSpacingMultiple: 1.5 });
	button(s, 'Learn More', 7.280, 5.318, 1.731, 0.443, ORANGE, 7.474, 5.388, 1.156);
}

/* 3 — market trends (dark) */
function slide03(p) {
	const s = p.addSlide();
	s.background = { color: DARK };
	chrome(s, true);
	// split photo: tall block + two stacked blocks
	photo(s, 6.888, 1.000, 2.657, 5.500);
	photo(s, 9.735, 1.000, 2.922, 2.667);
	photo(s, 9.744, 3.857, 2.913, 2.643);

	txt(s, 'Current Real Estate Market Trends', 0.676, 1.109, 5.046, 1.447,
		{ fontFace: TITLE, fontSize: 40, color: OFFWHITE });
	txt(s, 'Lorem ipsum dolor sit, consectetur elit, sed do  adipisicing sed  eiusmod incid. Maecenas porttitor congue massa. ',
		0.676, 2.671, 5.046, 0.630, { color: OFFWHITE, lineSpacingMultiple: 1.5 });

	shape(s, 'rect', 5.796, 4.512, 2.651, 1.833, { fill: { color: ORANGE }, shadow: DROP });
	[
		{ x: 0.676, value: '30% ' },
		{ x: 3.203, value: '500+' },
		{ x: 6.197, value: '85k' },
	].forEach(c => {
		txt(s, c.value, c.x, 4.592, 1.082, 0.650,
			{ fontFace: BODY_M, fontSize: 24, color: WHITE, lineSpacingMultiple: 1.5 });
		txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit.', c.x, 5.206, 1.847, 0.908,
			{ color: WHITE, lineSpacingMultiple: 1.5 });
	});
}

/* 4 — sectors */
function slide04(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	txt(s, 'Different Real Estate Sectors', 8.359, 1.178, 5.046, 1.447,
		{ fontFace: TITLE, fontSize: 40, color: BLACK });
	txt(s, LOREM_SECTOR, 8.359, 2.740, 4.212, 0.908, { lineSpacingMultiple: 1.5 });

	[
		{ y: 1.360, label: 'Residential', color: ICON_GREY },
		{ y: 3.296, label: 'Commercial', color: ORANGE },
		{ y: 5.233, label: 'Industrial', color: ICON_GREY },
	].forEach(r => {
		icon(s, 1.003, r.y + 0.128, 0.650, 0.650, r.color);
		txt(s, r.label, 1.944, r.y, 3.090, 0.303, { fontFace: BODY_M, fontSize: 12, color: DARK });
		txt(s, LOREM_SHORT, 1.944, r.y + 0.229, 3.090, 0.678, { fontSize: 12, lineSpacingMultiple: 1.5 });
	});

	photo(s, 5.613, 1.000, 2.107, 5.500);
	photo(s, 9.474, 5.032, 3.192, 1.468);
}

/* 5 — investments */
function slide05(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	txt(s, 'Profitable Real Estate Investments', 0.795, 1.140, 5.046, 1.447,
		{ fontFace: TITLE, fontSize: 40, color: BLACK });
	txt(s, LOREM_SECTOR, 0.795, 2.702, 5.046, 0.908, { lineSpacingMultiple: 1.5 });
	photo(s, 6.667, 1.000, 6.000, 2.750);

	shape(s, 'rect', 3.800, 4.490, 2.651, 2.010, { fill: { color: ORANGE }, shadow: GLOW });
	[
		{ x: 0.843, ix: 0.966, label: 'Rental Properties', on: false },
		{ x: 4.085, ix: 4.208, label: 'Vacation Rentals', on: true },
		{ x: 7.308, ix: 7.431, label: 'Commercial Real Estate ', on: false },
		{ x: 10.359, ix: 10.482, label: 'Commercial Real Estate ', on: false },
	].forEach(c => {
		icon(s, c.ix, 4.659, 0.600, 0.600, c.on ? DEEP_ORANGE : ICON_GREY);
		txt(s, c.label, c.x, 5.435, 2.242, 0.303,
			{ fontFace: BODY_M, fontSize: 12, color: c.on ? WHITE : DARK });
		txt(s, 'Lorem ipsum dolor sit, elit consectet adipisicing.', c.x, 5.665, 2.242, 0.630,
			{ color: c.on ? OFFWHITE : GREY, lineSpacingMultiple: 1.5 });
	});
}

/* 6 — risks (dark) */
function slide06(p) {
	const s = p.addSlide();
	s.background = { color: DARK };
	chrome(s, true);
	photo(s, 0.667, 1.000, 4.063, 5.500);
	txt(s, 'Risks in Real Estate Investments', 5.491, 1.196, 5.064, 1.447,
		{ fontFace: TITLE, fontSize: 40, color: OFFWHITE });
	txt(s, 'Lorem ipsum dolor sit, consectetur elit, sed do  adipisic sed  eiusmod incid. Maecenas porttitor congue massa. Maecenas porttitor congue massa. Fusce sed pulvinar ultricies, purus lectus.',
		5.491, 2.758, 5.731, 0.908, { color: OFFWHITE, lineSpacingMultiple: 1.5 });

	shape(s, 'rect', 5.491, 4.361, 3.451, 1.943, { fill: { color: ORANGE }, shadow: DROP });
	[
		{ x: 5.734, ix: 7.745, label: 'Challenges 01', color: DEEP_ORANGE },
		{ x: 9.459, ix: 11.597, label: 'Challenges 02', color: '443A37' },
	].forEach(c => {
		icon(s, c.ix, 4.543, 0.600, 0.600, c.color);
		txt(s, c.label, c.x, 5.096, 3.090, 0.303, { fontFace: BODY_M, fontSize: 12, color: WHITE });
		txt(s, LOREM_SHORT, c.x, 5.325, 3.090, 0.678,
			{ fontSize: 12, color: OFFWHITE, lineSpacingMultiple: 1.5 });
	});
}

/* 7 — emerging trends */
function slide07(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	txt(s, 'Emerging Trends in Real Estate', 0.795, 1.319, 4.888, 1.447,
		{ fontFace: TITLE, fontSize: 40, color: BLACK });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Mae porttitor congue massa. Fusce posuere, magna sed.',
		0.795, 2.882, 5.046, 0.630, { lineSpacingMultiple: 1.5 });
	photo(s, 0.000, 3.927, 6.175, 2.573);

	[
		{ y: 1.539, num: '01.', color: DARK, label: 'Sustainable & Green Buildings' },
		{ y: 3.296, num: '02.', color: ORANGE, label: 'Urbanization & Smart Cities' },
		{ y: 5.053, num: '03.', color: DARK, label: 'Shifts in Consumer Preferences' },
	].forEach(r => {
		txt(s, r.num, 7.239, r.y + 0.168, 1.100, 0.572, { fontFace: BODY_B, fontSize: 28, color: r.color });
		txt(s, r.label, 8.371, r.y, 3.959, 0.303, { fontFace: BODY_M, fontSize: 12, color: DARK });
		txt(s, LOREM_SHORT + ' sed pulvinar.', 8.371, r.y + 0.230, 3.959, 0.678,
			{ fontSize: 12, lineSpacingMultiple: 1.5 });
	});
}

/* 8 — specialists */
function slide08(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false, false);
	photo(s, 0.667, 3.397, 2.619, 4.103);
	photo(s, 3.750, 3.397, 2.619, 4.103);
	photo(s, 8.280, 1.000, 2.039, 2.000);
	photo(s, 10.492, 1.000, 2.175, 2.000);

	txt(s, 'Our Property Specialists', 8.312, 4.103, 3.856, 1.447,
		{ fontFace: TITLE, fontSize: 40, color: BLACK });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Mae porttitor congue.',
		8.312, 5.665, 3.628, 0.630, { lineSpacingMultiple: 1.5 });

	[
		{ x: 1.009, name: 'Mark Kennedy', role: 'CEO & Founder' },
		{ x: 4.092, name: 'Leon Choi', role: 'Head of Sales ' },
	].forEach(c => {
		txt(s, c.name, c.x, 1.565, 1.850, 0.337, { fontFace: BODY_M, fontSize: 14, color: DARK });
		txt(s, c.role, c.x, 1.834, 1.933, 0.375, { fontSize: 12, lineSpacingMultiple: 1.5 });
		txt(s, '\u201CLorem ipsum dolor sit, consectetur adipis\u201D', c.x, 2.139, 1.933, 0.630,
			{ lineSpacingMultiple: 1.5 });
	});
}

/* 9 — infographic: quartered circle (dark) */
function slide09(p) {
	const s = p.addSlide();
	s.background = { color: DARK };
	chrome(s, true);
	sectionTitle(s, true);

	// circle quarters: bounding box of the whole disc
	const CX = 6.667, CY = 4.140, R = 1.816;
	[
		{ from: 180, to: 270, color: WHITE },   // top-left
		{ from: 270, to: 360, color: ORANGE },  // top-right
		{ from: 0, to: 90, color: WHITE },      // bottom-right
		{ from: 90, to: 180, color: ORANGE },   // bottom-left
	].forEach(q => {
		shape(s, 'pie', CX - R, CY - R, R * 2, R * 2, {
			angleRange: [q.from, q.to],
			fill: { color: DARK }, line: { color: q.color, width: 1.75 },
		});
	});
	// white diamond plate + signpost icon
	shape(s, 'roundRect', 5.909, 3.382, 1.516, 1.516,
		{ rotate: 315.75, rectRadius: 0.08, fill: { color: WHITE }, shadow: SOFT });
	icon(s, 6.450, 3.783, 0.434, 0.714, DARK);
	// quadrant icons
	icon(s, 5.560, 3.048, 0.374, 0.345, WHITE);
	icon(s, 7.409, 3.047, 0.355, 0.346, WHITE);
	icon(s, 5.554, 4.895, 0.387, 0.329, WHITE);
	icon(s, 7.418, 4.891, 0.337, 0.338, WHITE);

	[
		{ x: 1.370, y: 1.912, label: 'Data Analysis 01', color: WHITE },
		{ x: 8.784, y: 1.912, label: 'Data Analysis 02', color: ORANGE },
		{ x: 8.784, y: 5.001, label: 'Data Analysis 03', color: WHITE },
		{ x: 1.370, y: 5.001, label: 'Data Analysis 04', color: ORANGE },
	].forEach(c => {
		shape(s, 'rect', c.x, c.y, 3.179, 1.409,
			{ fill: { color: DARK }, line: { color: c.color, width: 1.75 }, shadow: SOFT });
		txt(s, c.label, c.x + 0.478, c.y + 0.151, 2.224, 0.337,
			{ fontFace: BODY_M, fontSize: 14, color: WHITE, align: 'center' });
		txt(s, LOREM_MED, c.x + 0.185, c.y + 0.488, 2.809, 0.678,
			{ fontSize: 12, color: OFFWHITE, align: 'center', lineSpacingMultiple: 1.5 });
	});

	// elbow arrows: from the disc edge, step vertically, then into the card
	polyline(s, [[5.453, 2.793], [4.993, 2.793], [4.993, 2.620], [4.533, 2.620]], { color: WHITE, endCap: 'triangle' });
	polyline(s, [[7.880, 2.793], [8.307, 2.793], [8.307, 2.620], [8.787, 2.620]], { color: ORANGE, endCap: 'triangle' });
	polyline(s, [[7.847, 5.493], [8.293, 5.493], [8.293, 5.707], [8.787, 5.707]], { color: WHITE, endCap: 'triangle' });
	polyline(s, [[5.467, 5.493], [4.993, 5.493], [4.993, 5.707], [4.533, 5.707]], { color: ORANGE, endCap: 'triangle' });
}

/* 10 — infographic: six planning-stage cards */
function slide10(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	sectionTitle(s, false);

	const COLS = [1.766, 5.099, 8.432];
	const ROWS = [2.226, 4.268];
	const NUMS = [['01', '03', '05'], ['02', '04', '06']];
	ROWS.forEach((y, r) => COLS.forEach((x, c) => {
		const orange = (r + c) % 2 === 1;
		shape(s, 'rect', x, y, 3.135, 1.844, orange
			? { fill: { color: ORANGE }, shadow: DROP }
			: { fill: { color: DARK } });
		txt(s, NUMS[r][c], x + 0.384, y + 0.206, 0.766, 0.589,
			{ fontFace: BODY_M, fontSize: 32, bold: true, color: WHITE, margin: TIGHT });
		txt(s, 'PLANNING STAGE', x + 0.384, y + 0.847, 2.369, 0.286,
			{ fontSize: 14, color: WHITE, margin: TIGHT });
		txt(s, 'Mauris quam dolor, cursus at porta et, luctus eget', x + 0.384, y + 1.082, 2.369, 0.580,
			{ color: WHITE, lineSpacingMultiple: 1.5, margin: TIGHT });
	}));
}

/* 11 — infographic: eight numbered bars around a circle */
function slide11(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	sectionTitle(s, false);

	const ROWS = [1.996, 3.102, 4.208, 5.314];
	ROWS.forEach((y, i) => {
		// left column (odd numbers, right-aligned copy)
		const lc = i % 2 === 0 ? DARK : ORANGE;
		shape(s, 'rect', 1.077, y, 3.179, 0.948, { fill: { color: WHITE }, shadow: SOFT });
		shape(s, 'rect', 3.765, y + 0.083, 0.958, 0.782, { fill: { color: lc }, shadow: SOFT });
		txt(s, 'Lorem ipsum dolor sit, consectetur adipiscing', 1.336, y + 0.120, 2.171, 0.678,
			{ fontSize: 12, align: 'right', lineSpacingMultiple: 1.5 });
		txt(s, '0' + (i * 2 + 1), 3.944, y + 0.255, 0.599, 0.438,
			{ fontFace: BODY_M, fontSize: 20, bold: true, color: WHITE, align: 'center' });

		// right column (even numbers, mirrored)
		const rc = i % 2 === 0 ? ORANGE : DARK;
		shape(s, 'rect', 9.077, y, 3.179, 0.948, { fill: { color: WHITE }, flipH: true, shadow: SOFT });
		shape(s, 'rect', 8.611, y + 0.083, 0.958, 0.782, { fill: { color: rc }, flipH: true, shadow: SOFT });
		txt(s, 'Lorem ipsum dolor sit, consectetur adipiscing', 9.827, y + 0.120, 2.171, 0.678,
			{ fontSize: 12, lineSpacingMultiple: 1.5 });
		txt(s, '0' + (i * 2 + 2), 8.790, y + 0.255, 0.599, 0.438,
			{ fontFace: BODY_M, fontSize: 20, bold: true, color: WHITE, align: 'center' });
	});

	// hub connectors: bent runs to rows 01/02/07/08, straight ones to 03-06
	polyline(s, [[4.723, 2.861], [4.723, 3.211], [5.750, 3.211]], { color: DARK, endCap: 'triangle' });
	polyline(s, [[7.585, 3.211], [8.220, 3.211], [8.220, 2.469], [8.611, 2.469]], { color: ORANGE, endCap: 'triangle' });
	polyline(s, [[5.522, 3.576], [4.723, 3.576]], { color: ORANGE, endCap: 'triangle' });
	polyline(s, [[7.827, 3.576], [8.611, 3.576]], { color: DARK, endCap: 'triangle' });
	polyline(s, [[5.522, 4.682], [4.723, 4.682]], { color: DARK, endCap: 'triangle' });
	polyline(s, [[7.827, 4.682], [8.611, 4.682]], { color: ORANGE, endCap: 'triangle' });
	polyline(s, [[5.750, 5.046], [5.127, 5.046], [5.127, 5.788], [4.723, 5.788]], { color: ORANGE, endCap: 'triangle' });
	polyline(s, [[7.585, 5.046], [8.580, 5.046], [8.580, 5.788]], { color: DARK, endCap: 'triangle' });

	// hub
	shape(s, 'ellipse', 5.369, 2.832, 2.595, 2.595, { fill: { color: WHITE }, shadow: SOFT });
	shape(s, 'ellipse', 5.509, 2.971, 2.316, 2.316, { fill: { color: WHITE }, shadow: SOFT });
	icon(s, 6.070, 3.626, 1.193, 1.007, DARK);
}

/* 12 — infographic: nested diamonds (dark) */
function slide12(p) {
	const s = p.addSlide();
	s.background = { color: DARK };
	chrome(s, true);
	sectionTitle(s, true);

	// four large rotated squares (outer diamond ring)
	[
		{ x: 4.684, y: 3.324, color: WHITE },
		{ x: 5.843, y: 2.135, color: ORANGE },
		{ x: 5.872, y: 4.483, color: ORANGE },
		{ x: 7.032, y: 3.295, color: WHITE },
	].forEach(d => shape(s, 'rect', d.x, d.y, 1.618, 1.618, {
		rotate: 44.29, fill: { color: DARK }, line: { color: d.color, width: 1.75 }, shadow: SOFT,
	}));
	// four small rotated squares (inner ring)
	[
		{ x: 5.694, y: 3.750, w: 0.788, h: 0.767, color: ORANGE },
		{ x: 6.273, y: 3.155, w: 0.788, h: 0.767, color: WHITE },
		{ x: 6.288, y: 4.308, w: 0.788, h: 0.808, color: WHITE },
		{ x: 6.868, y: 3.714, w: 0.788, h: 0.808, color: ORANGE },
	].forEach(d => shape(s, 'rect', d.x, d.y, d.w, d.h, {
		rotate: 134.29, fill: { color: DARK }, line: { color: d.color, width: 1.75 }, shadow: SOFT,
	}));
	shape(s, 'ellipse', 6.180, 3.632, 0.973, 0.973, { fill: { color: WHITE }, shadow: SOFT });

	// satellite icons + centre trophy
	icon(s, 6.505, 2.338, 0.294, 0.379, WHITE);
	icon(s, 4.876, 3.989, 0.364, 0.287, WHITE);
	icon(s, 8.153, 3.932, 0.353, 0.343, WHITE);
	icon(s, 6.503, 5.578, 0.326, 0.304, WHITE);
	icon(s, 6.483, 3.930, 0.368, 0.377, DARK);

	[
		{ x: 1.200, y: 2.162, label: 'Project 02', color: WHITE },
		{ x: 9.446, y: 2.162, label: 'Project 03', color: ORANGE },
		{ x: 1.200, y: 4.672, label: 'Project 01', color: ORANGE },
		{ x: 9.446, y: 4.672, label: 'Project 04', color: WHITE },
	].forEach(c => {
		shape(s, 'rect', c.x, c.y, 2.688, 1.402,
			{ fill: { color: DARK }, line: { color: c.color, width: 1.75 }, shadow: SOFT });
		txt(s, c.label, c.x + 0.159, c.y + 0.207, 1.331, 0.337,
			{ fontSize: 14, bold: true, color: WHITE });
		txt(s, 'Lorem ipsum dolor sit amet, consectetur maecenas', c.x + 0.160, c.y + 0.521, 2.368, 0.630,
			{ color: WHITE, lineSpacingMultiple: 1.5 });
	});
}

/* 13 — infographic: four timeline tiles */
function slide13(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	sectionTitle(s, false);

	[
		{ tile: [1.371, 2.314], fill: DARK, ic: [1.688, 2.680, 0.886, 0.787], tx: 3.203, ty: 2.242, label: 'First Timeline', lc: '262626', years: '2019 - 2020', yc: DARK },
		{ tile: [6.667, 2.314], fill: ORANGE, ic: [7.061, 2.587, 0.730, 0.974], tx: 8.498, ty: 2.242, label: 'Second Timeline', lc: DARK, years: '2021 - 2022', yc: ORANGE },
		{ tile: [2.238, 4.604], fill: ORANGE, ic: [2.576, 4.942, 0.842, 0.842], tx: 4.069, ty: 4.532, label: 'Third Timeline', lc: DARK, years: '2023 - 2024', yc: ORANGE },
		{ tile: [7.533, 4.604], fill: DARK, ic: [7.819, 4.942, 0.948, 0.842], tx: 9.365, ty: 4.532, label: 'Fours Timeline', lc: '262626', years: '2025 - 2026', yc: DARK },
	].forEach(t => {
		shape(s, 'rect', t.tile[0], t.tile[1], 1.520, 1.520, t.fill === ORANGE
			? { fill: { color: ORANGE }, shadow: DROP }
			: { fill: { color: DARK } });
		icon(s, t.ic[0], t.ic[1], t.ic[2], t.ic[3], WHITE);
		txt(s, t.label, t.tx, t.ty, 2.939, 0.337, { fontFace: BODY_M, fontSize: 14, color: t.lc });
		txt(s, t.years, t.tx, t.ty + 0.400, 2.939, 0.572, { fontFace: BODY_M, fontSize: 28, color: t.yc });
		txt(s, 'Lorem ipsum dolor sit amet, elit, sed do eiusmod tempor', t.tx, t.ty + 0.981, 2.939, 0.707,
			{ fontSize: 12, lineSpacingMultiple: 1.5 });
	});
}

/* 14 — infographic: four diamonds on a zig-zag */
function slide14(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	sectionTitle(s, false);

	const CAPTION = [
		{ text: 'Lorem ipsum dolor sit amet, consecte', options: { breakLine: true } },
		{ text: 'tur adipiscing elit, ' },
	];
	[
		{ dx: 2.321, dy: 2.115, fill: DARK, tx: 2.274, ty: 4.216, cx: 2.291, year: '2009', ic: [3.014, 2.860, 0.478, 0.343] },
		{ dx: 4.680, dy: 2.708, fill: ORANGE, tx: 4.648, ty: 4.809, cx: 4.651, year: '2011', ic: [5.419, 3.356, 0.346, 0.488] },
		{ dx: 6.858, dy: 2.115, fill: DARK, tx: 6.778, ty: 4.216, cx: 6.795, year: '2013', ic: [7.635, 2.832, 0.315, 0.397] },
		{ dx: 9.081, dy: 2.708, fill: ORANGE, tx: 9.095, ty: 4.809, cx: 9.112, year: '2015', ic: [9.802, 3.411, 0.429, 0.450] },
	].forEach(d => {
		shape(s, 'diamond', d.dx, d.dy, 1.871, 1.871, d.fill === ORANGE
			? { fill: { color: ORANGE }, shadow: DROP }
			: { fill: { color: DARK } });
		icon(s, d.ic[0], d.ic[1], d.ic[2], d.ic[3], WHITE);
		txt(s, d.year, d.tx, d.ty, 1.964, 0.438, { fontFace: BODY_M, fontSize: 20, color: DARK, align: 'center' });
		txt(s, CAPTION, d.cx, d.ty + 0.481, 1.918, 0.978,
			{ fontSize: 12, align: 'center', lineSpacingMultiple: 1.5 });
	});
}

/* 15 — infographic: arrow bar chart (dark) */
function slide15(p) {
	const s = p.addSlide();
	s.background = { color: DARK };
	chrome(s, true);
	sectionTitle(s, true);

	// left-hand month notes, two columns of three
	[
		{ x: 1.324, months: ['January', 'March', 'May'] },
		{ x: 3.764, months: ['February', 'April', 'June'] },
	].forEach(col => col.months.forEach((m, i) => {
		const y = 2.490 + i * 1.0855;
		txt(s, m, col.x, y, 1.147, 0.337, { fontFace: BODY_M, fontSize: 14, color: WHITE });
		txt(s, 'Lorem ipsum dolor sit ut amet, elit. Sed', col.x, y + 0.277, 2.153, 0.678,
			{ fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5 });
	}));

	// baseline
	shape(s, 'line', 6.629, 5.800, 5.381, 0.000, { line: { color: RULE, width: 1 } });

	// bars are homePlate arrows rotated to point up; x/y/w/h are pre-rotation
	[
		{ bx: 5.723, by: 3.845, bw: 3.210, color: WHITE, label: 'Jan', pct: '80%', px: 7.026, py: 2.197 },
		{ bx: 6.742, by: 4.064, bw: 2.765, color: ORANGE, label: 'Feb', pct: '68%', px: 7.822, py: 2.642 },
		{ bx: 7.391, by: 3.917, bw: 3.060, color: WHITE, label: 'Mar', pct: '74%', px: 8.619, py: 2.347 },
		{ bx: 7.965, by: 3.695, bw: 3.504, color: ORANGE, label: 'Apr', pct: '97%', px: 9.415, py: 1.836 },
		{ bx: 9.343, by: 4.276, bw: 2.341, color: WHITE, label: 'May', pct: '52%', px: 10.211, py: 3.067 },
		{ bx: 9.558, by: 3.695, bw: 3.504, color: ORANGE, label: 'Jun', pct: '97%', px: 11.008, py: 1.836 },
	].forEach(b => {
		shape(s, 'homePlate', b.bx, b.by, b.bw, 0.707, {
			rotate: 270, fill: { color: DARK }, line: { color: b.color, width: 1.75 }, shadow: SOFT,
		});
		txt(s, b.pct, b.px, b.py, 0.605, 0.375,
			{ fontSize: 12, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
		txt(s, b.label, b.px, 5.827, 0.605, 0.375,
			{ fontSize: 12, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
	});
}

/* 16 — infographic: funnel */
function slide16(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	sectionTitle(s, false);

	[
		{ x: 1.893, w: 2.898, y: 2.027, value: '2,536 k', fill: DARK },
		{ x: 1.539, w: 3.606, y: 3.077, value: '11,489 k', fill: ORANGE },
		{ x: 1.204, w: 4.275, y: 4.126, value: '38,338 k', fill: DARK },
		{ x: 0.850, w: 4.983, y: 5.175, value: '69,782 k', fill: ORANGE },
	].forEach(b => {
		shape(s, 'trapezoid', b.x, b.y, b.w, 0.888, Object.assign(
			{ flipH: true, fill: { color: b.fill } },
			b.fill === ORANGE ? { shadow: DROP } : { shadow: SOFT }));
		txt(s, b.value, 2.482, b.y + 0.226, 1.719, 0.438,
			{ fontFace: BODY_M, fontSize: 20, color: WHITE, align: 'center' });
	});

	[
		{ x: 6.964, y: 2.481, num: '01' },
		{ x: 9.715, y: 2.481, num: '02' },
		{ x: 6.964, y: 4.259, num: '03' },
		{ x: 9.715, y: 4.259, num: '04' },
	].forEach(k => {
		txt(s, k.num, k.x, k.y, 0.792, 0.438, { fontFace: BODY_M, fontSize: 20, color: DARK });
		txt(s, 'Keyword Here', k.x + 0.001, k.y + 0.388, 1.672, 0.337,
			{ fontFace: BODY_M, fontSize: 14, color: DARK });
		txt(s, LOREM_MED, k.x + 0.001, k.y + 0.676, 2.767, 0.630, { lineSpacingMultiple: 1.5 });
	});
}

/* 17 — infographic: five-step fan */
function slide17(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	sectionTitle(s, false);

	// wedges share an apex; radius shrinks slightly from step 01 to step 05
	const APEX_X = 6.828, APEX_Y = 6.232;
	[
		{ from: 180, to: 216, r: 3.319, fill: DARK },
		{ from: 216, to: 252, r: 3.100, fill: ORANGE },
		{ from: 252, to: 288, r: 2.859, fill: DARK },
		{ from: 288, to: 324, r: 2.617, fill: ORANGE },
		{ from: 324, to: 360, r: 2.368, fill: DARK },
	].forEach(w => shape(s, 'pie', APEX_X - w.r, APEX_Y - w.r, w.r * 2, w.r * 2, {
		angleRange: [w.from, w.to], fill: { color: w.fill },
		shadow: w.fill === ORANGE ? DROP : SOFT,
	}));

	// step captions sit inside the wedges (number over the word STEP)
	[
		{ x: 4.806, y: 5.361, w: 0.731, num: '01' },
		{ x: 5.332, y: 4.353, w: 0.749, num: '02' },
		{ x: 6.413, y: 4.070, w: 0.756, num: '03' },
		{ x: 7.485, y: 4.353, w: 0.780, num: '04' },
		{ x: 8.022, y: 5.361, w: 0.777, num: '05' },
	].forEach(st => txt(s, [
		{ text: st.num, options: { fontSize: 32, breakLine: true } },
		{ text: 'STEP', options: { fontSize: 14 } },
	], st.x, st.y, st.w, 0.728, {
		fontFace: BODY_M, color: WHITE, align: 'center', wrap: false, lineSpacingMultiple: 0.8,
	}));

	// dashed leader lines, each ending in a small dot next to its caption
	const leader = { color: RULE, dash: 'dash', startCap: 'oval' };
	polyline(s, [[3.601, 5.410], [3.266, 5.410]], { color: RULE, dash: 'dash', endCap: 'oval' });
	polyline(s, [[4.822, 3.886], [4.822, 3.577], [3.836, 3.577]], leader);
	polyline(s, [[6.775, 3.396], [6.775, 2.467], [5.789, 2.467]], leader);
	polyline(s, [[8.292, 4.093], [8.292, 3.577], [9.278, 3.577]], leader);
	polyline(s, [[9.017, 5.410], [9.674, 5.410]], { color: RULE, dash: 'dash', endCap: 'oval' });

	[
		{ tx: 1.167, ty: 3.299, bx: 1.161, by: 3.695 },
		{ tx: 0.687, ty: 4.857, bx: 0.664, by: 5.252 },
		{ tx: 3.071, ty: 2.036, bx: 3.048, by: 2.432 },
		{ tx: 9.695, ty: 3.255, bx: 9.869, by: 3.690 },
		{ tx: 10.138, ty: 4.857, bx: 10.133, by: 5.282 },
	].forEach(l => {
		txt(s, 'Add title here', l.tx, l.ty, 1.713, 0.370,
			{ fontFace: BODY_M, fontSize: 16, wrap: false });
		txt(s, 'Sed perspiciatis unde omnis  elit voluptatem fringilla.', l.bx, l.by, 2.581, 0.678,
			{ fontSize: 12, lineSpacingMultiple: 1.5 });
	});
}

/* 18 — infographic: five arrow callouts (dark) */
function slide18(p) {
	const s = p.addSlide();
	s.background = { color: DARK };
	chrome(s, true);
	sectionTitle(s, true);

	[
		{ x: 1.135, up: true, color: ORANGE, label: 'Performance 01', lw: 1.878, ix: 1.974, iy: 4.114, iw: 0.370, ih: 0.369 },
		{ x: 3.389, up: false, color: WHITE, label: 'Performance 02', lw: 1.820, ix: 4.245, iy: 4.138, iw: 0.335, ih: 0.345 },
		{ x: 5.643, up: true, color: ORANGE, label: 'Performance 03', lw: 1.878, ix: 6.482, iy: 4.143, iw: 0.369, ih: 0.341 },
		{ x: 7.897, up: false, color: WHITE, label: 'Performance 04', lw: 1.816, ix: 8.704, iy: 4.173, iw: 0.433, ih: 0.310 },
		{ x: 10.151, up: true, color: ORANGE, label: 'Performance 05', lw: 1.878, ix: 11.006, iy: 4.146, iw: 0.337, ih: 0.338 },
	].forEach(c => {
		shape(s, 'upArrowCallout', c.x, c.up ? 3.517 : 3.933, 2.048, 1.161, {
			flipH: !c.up, flipV: !c.up,
			fill: { color: DARK }, line: { color: c.color, width: 1.75 }, shadow: SOFT,
		});
		icon(s, c.ix, c.iy, c.iw, c.ih, WHITE);
		const ty = c.up ? 2.339 : 5.324;
		txt(s, c.label, c.x + (c.up ? 0.085 : 0.114), ty, c.lw, 0.337,
			{ fontFace: BODY_M, fontSize: 14, color: WHITE, align: 'center' });
		txt(s, 'Lorem ipsum dolor sit amet, consectetur', c.x - 0.085, ty + 0.274, 2.217, 0.678,
			{ fontSize: 12, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
	});
}

/* 19 — contact */
function slide19(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false);
	photo(s, 7.048, 1.000, 5.618, 5.500);
	txt(s, 'Reach Out to Us', 0.843, 1.467, 4.888, 0.774, { fontFace: TITLE, fontSize: 40, color: BLACK });
	txt(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Mae porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero.',
		0.843, 2.394, 5.046, 0.908, { lineSpacingMultiple: 1.5 });

	photo(s, 0.667, 4.087, 4.349, 2.021);
	shape(s, 'rect', 5.511, 4.087, 4.349, 2.021, { fill: { color: ORANGE }, shadow: GLOW });
	[
		{ y: 4.235, iy: 4.451, label: 'Location', value: 'Your Business Location' },
		{ y: 5.132, iy: 5.343, label: 'Email', value: 'businessemail@gmail.com' },
	].forEach(r => {
		icon(s, 6.085, r.iy, 0.409, 0.409, DEEP_ORANGE);
		txt(s, r.label, 6.757, r.y, 2.108, 0.421, { fontFace: BODY_M, fontSize: 14, color: LIGHT });
		txt(s, r.value, 6.768, r.y + 0.346, 2.520, 0.375,
			{ fontSize: 12, color: WHITE, lineSpacingMultiple: 1.5 });
	});
}

/* 20 — thank you */
function slide20(p) {
	const s = p.addSlide();
	s.background = { color: LIGHT };
	chrome(s, false, false);
	photo(s, 0.000, 1.000, 5.860, 6.500);
	[1.000, 1.690, 2.379].forEach(y => photo(s, 8.309, y, 5.018, 0.630));

	shape(s, 'rect', 4.653, 1.368, 1.647, 0.774, { fill: { color: DARK }, shadow: DROP });
	txt(s, 'Feel Free to Contact Us!', 4.839, 1.503, 1.277, 0.505,
		{ fontFace: BODY_M, fontSize: 12, color: WHITE, align: 'center' });
	button(s, 'End Presentation', 4.653, 2.336, 2.100, 0.443, ORANGE, 4.791, 2.406, 1.615);

	txt(s, 'Thank You for Attention', 8.309, 3.943, 3.969, 1.582,
		{ fontFace: TITLE, fontSize: 44, color: BLACK });
	txt(s, '\u201CLorem ipsum dolor sit, consectetur adipis Sed perspiciatis unde omnis  elit voluptatem fringilla\u201D',
		8.309, 5.607, 4.258, 0.630, { italic: true, lineSpacingMultiple: 1.5 });
}

/* ------------------------------------------------------------------ */
/* build                                                               */
/* ------------------------------------------------------------------ */
function build() {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'WIDE_16x9', width: 13.333, height: 7.5 });
	pptx.layout = 'WIDE_16x9';
	pptx.author = 'Horizon Estates';
	pptx.title = 'Real Estate Presentation';

	[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
		slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
		.forEach(fn => fn(pptx));

	const out = path.join(__dirname, '0526d5a7-fa39-4a29-aad8-9994dce0c51e_grok_final.pptx');
	return pptx.writeFile({ fileName: out }).then(() => console.log('wrote ' + out));
}

build().catch(err => { console.error(err); process.exit(1); });
