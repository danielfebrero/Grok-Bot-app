/**
 * Arazuna - Architecture & Property Agency Presentation
 * Standalone pptxgenjs recreation of the 34-slide reference deck (13.333in x 7.5in).
 *
 * Raster photographs in the reference are replaced by flat grey placeholder
 * rectangles (see `photo()`); every other element is a native pptxgenjs shape.
 */
const path = require('path');
const pptxgen = require('pptxgenjs');

// ---------------------------------------------------------------- palette --
const RED = 'FF2525';
const DARK = '222A35';
const WHITE = 'FFFFFF';
const GREY = '7F7F7F';
const GREY_D = '757070';
const SILVER = 'F2F2F2';
const PHOTO = 'C3C3C3';
const PHOTO_D = '7F7F7F';   // the deck's second, darker photo
const HAIRLINE = 'D8D8D8';
const DASHGREY = 'AEABAB';
const BLUE = '2E75B5';

// ------------------------------------------------------------------ fonts --
const MONT = 'Montserrat';
const MONT_BLACK = 'Montserrat Black';
const MONT_XB = 'Montserrat ExtraBold';
const MONT_SB = 'Montserrat SemiBold';
const POPPINS = 'Poppins SemiBold';
const LATO = 'Lato';

// ------------------------------------------------------- reusable strings --
const LOREM_L = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. Maecenas accumsan lacus vel facilisis volutpat est velit egestas dui. ';
const LOREM_M = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. ';
const LOREM_S = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus';
const LOREM_XS = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';
const LOREM_CARD = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.';
const LOREM_TINY = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed';
const QUOTE = 'Good buildings come from good people and all problem are solved by good design. ';
// pptxgenjs rewrites shadow objects in place while serialising, so every shape
// needs its own instance - these builders hand out a fresh one each call.
const softShadow = (color) => ({ type: 'outer', color, blur: 25, offset: 3, angle: 90, opacity: 0.1 });
const cardShadow = () => softShadow(DARK);
const redShadow = () => softShadow(RED);
const dropShadow = () => ({ type: 'outer', color: '000000', blur: 4.5, offset: 1.5, angle: 90, opacity: 0.63 });

// ----------------------------------------------------------- tiny helpers --
/** Text block. Defaults mirror the deck: top-anchored, left aligned Lato body. */
function t(s, text, o) {
	s.addText(text, Object.assign({ fontFace: LATO, fontSize: 10, color: GREY, align: 'left', valign: 'top' }, o));
}
/** Section headline (bold Montserrat, dark). */
function title(s, text, x, y, w, h, o) {
	t(s, text, Object.assign({ x, y, w, h, fontFace: MONT, fontSize: 28, bold: true, color: DARK }, o));
}
/** Small bold caption ("Your Title Here" style). */
function caption(s, text, x, y, w, o) {
	t(s, text, Object.assign({ x, y, w, h: 0.303, fontFace: MONT, fontSize: 12, bold: true, color: DARK }, o));
}
/** Grey lorem paragraph with the deck's 1.5 line spacing. */
function para(s, text, x, y, w, h, o) {
	t(s, text, Object.assign({ x, y, w, h, lineSpacingMultiple: 1.5 }, o));
}
function rect(s, x, y, w, h, color, o) {
	s.addShape('rect', Object.assign({ x, y, w, h, fill: { color } }, o));
}
/** Placeholder standing in for a photograph in the reference deck. */
function photo(s, x, y, w, h, o) {
	s.addShape('rect', Object.assign({ x, y, w, h, fill: { color: PHOTO } }, o));
}
/** Thin connector with a stealth arrow head, as used by all "read more" links. */
function arrow(s, x, y, w, color, dashType) {
	s.addShape('line', { x, y, w, h: 0, line: { color, width: 0.75, dashType: dashType || 'solid', endArrowType: 'stealth' } });
}
/** Arazuna mark: three nested triangles. `w` is the overall width. */
function logo(s, x, y, w) {
	const k = w / 0.396;
	s.addShape('triangle', { x, y, w, h: 0.341 * k, fill: { color: RED } });
	s.addShape('triangle', { x: x + 0.091 * k, y: y + 0.105 * k, w: 0.212 * k, h: 0.183 * k, fill: { color: WHITE } });
	s.addShape('triangle', { x: x + 0.113 * k, y: y + 0.208 * k, w: 0.168 * k, h: 0.145 * k, fill: { color: DARK }, line: { color: WHITE, width: 2.2 * k } });
}
/** Round "next" button: white disc with a dark arrow. */
function discArrow(s, x, y, d) {
	s.addShape('ellipse', { x, y, w: d, h: d, fill: { color: WHITE } });
	arrow(s, x + 0.24 * d, y + d / 2, 0.52 * d, DARK);
}
/** Red pill button with label + arrow ("Read More" / "Learn More"). */
function pillButton(s, label, x, y) {
	s.addShape('roundRect', { x, y, w: 2.188, h: 0.52, rectRadius: 0.087, fill: { color: RED } });
	caption(s, label, x + 0.11, y + 0.109, 1.217, { color: WHITE });
	arrow(s, x + 1.407, y + 0.26, 0.549, WHITE);
}
/** Hexagon badge: white outer hexagon + coloured inner hexagon + glyph. */
function hexBadge(s, cx, cy, color, glyphName) {
	s.addShape('hexagon', { x: cx - 0.7555, y: cy - 0.6765, w: 1.511, h: 1.353, rotate: 90, fill: { color: WHITE }, shadow: softShadow('000000') });
	s.addShape('hexagon', { x: cx - 0.6335, y: cy - 0.5675, w: 1.267, h: 1.135, rotate: 90, fill: { color } });
	glyph(s, glyphName, cx, cy, 0.56, WHITE);
}

/**
 * Simplified line-art icons standing in for the reference deck's icon groups.
 * Each icon is drawn from primitives on a 10x10 grid centred on (cx, cy) with side `sz`.
 */
function glyph(s, kind, cx, cy, sz, color) {
	const u = sz / 10;
	const at = (gx, gy) => ({ x: cx + (gx - 5) * u, y: cy + (gy - 5) * u });
	const stroke = { color, width: Math.max(0.5, sz * 1.4) };
	const box = (gx, gy, gw, gh) => s.addShape('rect', Object.assign(at(gx, gy), { w: gw * u, h: gh * u, fill: { type: 'none' }, line: stroke }));
	const bar = (gx, gy, gw, gh) => s.addShape('rect', Object.assign(at(gx, gy), { w: gw * u, h: gh * u, fill: { color } }));
	const ring = (gx, gy, gw, gh) => s.addShape('ellipse', Object.assign(at(gx, gy), { w: gw * u, h: gh * u, fill: { type: 'none' }, line: stroke }));
	const dot = (gx, gy, gd) => s.addShape('ellipse', Object.assign(at(gx, gy), { w: gd * u, h: gd * u, fill: { color } }));
	const dome = (gx, gy, gw, gh, filled) => s.addShape('pie', Object.assign(at(gx, gy), { w: gw * u, h: gh * u, angleRange: [180, 360], fill: filled ? { color } : { type: 'none' }, line: stroke }));
	const tri = (gx, gy, gw, gh, rot) => s.addShape('triangle', Object.assign(at(gx, gy), { w: gw * u, h: gh * u, rotate: rot || 0, fill: { color } }));

	switch (kind) {
		case 'building':                                   // office block, three window rows
			box(1.8, 1.5, 6.4, 8.5);
			[3, 5, 7].forEach((gy) => [2.8, 4.5, 6.2].forEach((gx) => bar(gx, gy, 1.1, 1.1)));
			bar(4.3, 8.6, 1.4, 1.4);
			break;
		case 'tower':                                      // two towers, uneven heights
			box(1.2, 3.6, 3.4, 6.4); box(5.2, 1.2, 3.6, 8.8);
			bar(2, 4.8, 1.8, 0.7); bar(6, 2.6, 2, 0.7); bar(6, 4.8, 2, 0.7); bar(6, 7, 2, 0.7);
			break;
		case 'house':                                      // roof + body + door
			tri(1.6, 1.2, 6.8, 4.2); box(2.6, 4.8, 4.8, 5.2); bar(4.2, 7.4, 1.6, 2.6);
			break;
		case 'plan':                                       // floor-plan grid
			box(1, 1.8, 8, 6.4); bar(4.4, 1.8, 0.28, 6.4); bar(1, 5.2, 8, 0.28); bar(4.4, 3.4, 4.6, 0.28);
			break;
		case 'ruler':                                      // set square with inner frame
			s.addShape('rtTriangle', Object.assign(at(1.2, 1.4), { w: 7.6 * u, h: 7.6 * u, flipH: true, fill: { type: 'none' }, line: stroke }));
			box(3.4, 4.2, 3.2, 3.2);
			break;
		case 'helmet':                                     // hard-hat worker: cap, brim, face, shoulders
			dome(3, 0.6, 4, 3, false); bar(1.6, 3.6, 6.8, 0.7); ring(3.4, 4.4, 3.2, 2.6); dome(1.6, 6.4, 6.8, 4.6, false);
			break;
		case 'clock':                                      // stopwatch
			ring(1, 2, 8, 8); bar(4.75, 4, 0.5, 2.2); bar(4.9, 5.7, 1.9, 0.5); bar(4.1, 0.6, 1.8, 1.2);
			break;
		case 'drop':
			tri(3, 0.6, 4, 3.8); dot(3, 3.2, 4);
			break;
		case 'star':
			s.addShape('star5', Object.assign(at(0, 0), { w: sz, h: sz, fill: { type: 'none' }, line: stroke }));
			break;
		case 'database':                                   // stacked discs
			[1.2, 4, 6.8].forEach((gy) => ring(1.4, gy, 7.2, 2.4));
			break;
		case 'gear':
			s.addShape('gear6', Object.assign(at(0.4, 0.4), { w: 9.2 * u, h: 9.2 * u, fill: { type: 'none' }, line: stroke }));
			ring(3.4, 3.4, 3.2, 3.2);
			break;
		case 'bulb':
			ring(2.4, 0.6, 5.2, 5.2); bar(4.2, 5.6, 1.6, 2.6); bar(3.8, 8.4, 2.4, 0.7);
			break;
		case 'chart':
			bar(1.6, 6, 1.8, 4); bar(4.1, 3.4, 1.8, 6.6); bar(6.6, 1.2, 1.8, 8.8);
			break;
		case 'note':                                       // notebook with spiral
			box(2.6, 1, 6.4, 8.6); [2.2, 4.5, 6.8].forEach((gy) => bar(1.2, gy, 2.2, 0.7));
			break;
		case 'send':                                       // paper plane
			tri(1.2, 1.6, 7.6, 6.8, 135);
			break;
		case 'target':
			ring(0.4, 0.4, 9.2, 9.2); ring(3, 3, 4, 4); dot(4.5, 4.5, 1);
			break;
		case 'handshake':
			bar(0.8, 4.2, 3.8, 1.8); bar(5.4, 4.2, 3.8, 1.8); dot(3.2, 3.4, 3.4);
			break;
		case 'growth':                                     // rising arrow over a baseline
			bar(0.8, 8.6, 8.4, 0.7); bar(1.2, 6.2, 6, 0.6); tri(6.8, 1, 2.6, 2.6, 45);
			break;
		case 'twitter':                                    // simplified bird
			bar(2, 4, 6, 2.4); tri(0.8, 2.8, 3, 3, 315);
			break;
		case 'facebook':                                   // f mark
			bar(3.8, 1.4, 2.6, 7.2); bar(2, 3.6, 6, 1.8);
			break;
		case 'instagram':                                  // rounded square + lens
			box(1.4, 1.4, 7.2, 7.2); ring(3.2, 3.2, 3.6, 3.6);
			break;
		default:
			ring(1.5, 1.5, 7, 7);
	}
}

// =========================================================== slide builders =
const slides = [];

// 1 — cover
slides.push((s) => {
	photo(s, 6.667, 0, 6.667, 6.356);
	t(s, 'Arazuna', { x: 1.047, y: 1.958, w: 4.589, h: 1.111, fontFace: MONT_BLACK, fontSize: 60, bold: true, color: DARK });
	t(s, 'Architecture & Property Agency Presentation', { x: 1.047, y: 3.068, w: 4.589, h: 0.32, fontSize: 13, color: GREY_D });
	para(s, 'Lorem ipsum dolor sit amet, lacus nulla ac netus nibh aliquet, porttitor ligula justo libero vivamus porttitor dolor, conubia mollit. Sapien nam suspendisse, tincidunt eget ante tincidunt, ', 1.047, 4.086, 4.589, 0.825);
	caption(s, 'Read More', 1.047, 5.399, 2.133);
	rect(s, 6.667, 5.734, 5.264, 1.245, RED);
	para(s, QUOTE, 7.276, 5.905, 2.937, 0.903, { fontFace: MONT, fontSize: 11, italic: true, color: WHITE });
	discArrow(s, 10.804, 6.039, 0.535);
	logo(s, 0.387, 0.334, 0.396);
});

// 2 — welcome
slides.push((s) => {
	photo(s, 0, 0, 6.203, 7.5);
	rect(s, 1.362, 6.471, 4.841, 1.029, RED);
	['Building ', 'Apartment', 'Housing'].forEach((label, i) => {
		caption(s, label, [1.744, 3.201, 4.802][i], 6.834, 1.437, { color: WHITE });
	});
	rect(s, 6.203, 6.471, 1.362, 1.029, DARK);
	discArrow(s, 6.663, 6.765, 0.442);
	logo(s, 12.386, 0.327, 0.396);
	title(s, 'Welcome to Arazuna Agency Architecture', 7.438, 1.768, 4.679, 1.043);
	para(s, LOREM_L, 7.438, 3.262, 4.679, 1.078);
	para(s, LOREM_M, 7.438, 4.461, 4.679, 0.825);
});

// 3 — history
slides.push((s) => {
	rect(s, 8.069, 0.678, 5.264, 6.822, DARK);
	title(s, 'History of Arazuna Agency Architecture', 0.981, 1.433, 4.679, 1.043);
	para(s, LOREM_L, 0.981, 2.927, 4.679, 1.078);
	para(s, LOREM_M, 0.981, 4.125, 4.679, 0.825);
	t(s, '1998', { x: 11.863, y: 2.739, w: 1.284, h: 0.37, fontFace: MONT, fontSize: 16, color: WHITE, align: 'center' });
	t(s, '1999', { x: 11.863, y: 3.499, w: 1.284, h: 0.505, fontFace: MONT, fontSize: 24, bold: true, color: RED, align: 'center' });
	t(s, '2000', { x: 11.853, y: 4.39, w: 1.284, h: 0.37, fontFace: MONT, fontSize: 16, color: WHITE, align: 'center' });
	[3.311, 4.216].forEach((y) => s.addShape('line', { x: 12.183, y, w: 0.623, h: 0, line: { color: HAIRLINE, width: 0.75 } }));
	pillButton(s, 'Read More  ', 1.087, 5.393);
	logo(s, 0.387, 0.334, 0.396);
	photo(s, 6.667, 0.678, 5.0, 6.144);
});

// 4 — stats banner
slides.push((s) => {
	photo(s, 0, 0, 13.333, 4.768);
	rect(s, 6.98, 4.011, 6.353, 1.515, RED);
	[['1170+', 'Coustomer', 7.319, 7.599, 1.348], ['12K', 'Project Success', 9.202, 9.651, 1.011], ['115', 'Winning Award', 11.098, 11.535, 1.011]]
		.forEach(([big, small, gx, bx, bw]) => {
			t(s, big, { x: bx, y: 4.358, w: bw, h: 0.505, fontFace: MONT, fontSize: 24, bold: true, color: WHITE, align: 'center' });
			t(s, small, { x: gx, y: 4.892, w: 1.909, h: 0.286, fontFace: MONT_SB, fontSize: 11, color: WHITE, align: 'center' });
		});
	[9.19, 11.073].forEach((x) => s.addShape('roundRect', { x, y: 4.452, w: 0.05, h: 0.632, rectRadius: 0.025, fill: { color: SILVER } }));
	title(s, 'We Are the Best Architecture Agency in America', 1.053, 5.306, 4.874, 1.515);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. Maecenas accumsan lacus vel facilisis', 6.98, 5.996, 5.687, 0.825);
	logo(s, 0.373, 0.29, 0.432);
});

// 5 — vision / mission
slides.push((s) => {
	rect(s, 0, 5.125, 13.333, 2.403, DARK);
	title(s, 'We Present a New Concept in Modern Architecture', 0.714, 1.313, 4.681, 1.515, { fontSize: 27 });
	para(s, LOREM_L, 0.714, 3.193, 4.681, 1.078);
	photo(s, 6.108, 0, 3.212, 5.125);
	photo(s, 9.705, 2.375, 3.212, 5.125);
	caption(s, 'New Concept', 9.705, 0.876, 1.909);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt', 9.705, 1.179, 3.212, 0.573);
	[['01. Vision', 0.714], ['02. Mission', 4.758]].forEach(([label, x]) => caption(s, label, x, 5.762, 1.909, { color: WHITE }));
	[0.714, 4.755].forEach((x) => para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor', x, 6.105, 3.328, 0.825, { color: WHITE }));
	logo(s, 0.387, 0.334, 0.396);
});

// 6 — project timeline
slides.push((s) => {
	title(s, 'Our Project Timeline', 4.364, 0.871, 4.681, 0.572, { align: 'center' });
	para(s, LOREM_S, 3.682, 1.713, 5.97, 0.573, { align: 'center' });
	const cols = [
		{ x: 1.755, icon: 'plan', color: RED, capX: 0.801, bodyX: 0.623 },
		{ x: 4.211, icon: 'house', color: RED, capX: 3.256, bodyX: 3.079 },
		{ x: 6.667, icon: 'tower', color: RED, capX: 5.712, bodyX: 5.535 },
		{ x: 9.122, icon: 'clock', color: RED, capX: 8.168, bodyX: 7.991 },
		{ x: 11.578, icon: 'building', color: DARK, capX: 10.623, bodyX: 10.446 },
	];
	cols.forEach((c, i) => {
		hexBadge(s, c.x, 4.0, c.color, c.icon);
		caption(s, 'Your Title Here', c.capX, 5.333, 1.909, { align: 'center' });
		para(s, LOREM_XS, c.bodyX, 5.773, 2.264, 0.825, { align: 'center' });
		if (i < 4) arrow(s, c.x + 0.952, 4.028, 0.549, DASHGREY, 'dash');
	});
	logo(s, 12.386, 0.327, 0.396);
});

// 7 — why choose us
slides.push((s) => {
	photo(s, 0.648, 0.539, 12.04, 4.1);
	rect(s, 0.647, 0.55, 12.04, 4.1, DARK, { fill: { color: DARK, transparency: 20 } });
	title(s, 'Why Choose Us?', 4.326, 1.852, 4.681, 0.606, { fontSize: 30, color: WHITE, align: 'center' });
	const cards = [
		{ x: 1.108, bg: RED, fg: WHITE, body: WHITE, icon: 'helmet', iconColor: WHITE },
		{ x: 4.926, bg: WHITE, fg: DARK, body: GREY, icon: 'house', iconColor: RED },
		{ x: 8.869, bg: WHITE, fg: DARK, body: GREY, icon: 'building', iconColor: RED },
	];
	cards.forEach((c) => {
		rect(s, c.x, 3.918, 3.482, 1.456, c.bg, { shadow: c.bg === RED ? redShadow() : cardShadow() });
		glyph(s, c.icon, c.x + 0.49, 4.42, 0.55, c.iconColor);
		caption(s, 'Your Title Here', c.x + 1.026, 4.081, 1.909, { color: c.fg });
		para(s, LOREM_XS, c.x + 1.026, 4.427, 2.264, 0.825, { color: c.body });
	});
	para(s, LOREM_S, 3.308, 6.102, 6.717, 0.573, { align: 'center' });
});

// 8 — price cards
slides.push((s) => {
	photo(s, 0, 1.387, 7.186, 4.726);
	[[3.226, RED, '$ 3.235.000', 3.465, 4.944, 6.573], [8.053, DARK, '$ 1.745.000', 8.285, 4.953, 11.394]]
		.forEach(([x, color, price, tx, ty, dx]) => {
			rect(s, x, 4.83, 3.959, 1.283, color, { shadow: color === RED ? redShadow() : cardShadow() });
			t(s, price, { x: tx, y: ty, w: 1.877, h: 0.404, fontFace: MONT, fontSize: 18, color: WHITE });
			t(s, 'consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et', { x: tx, y: ty + 0.452, w: 3.5, h: 0.561, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.3 });
			s.addShape('ellipse', { x: dx, y: ty, w: 0.412, h: 0.412, fill: { color: WHITE } });
			s.addShape('line', { x: dx + 0.115, y: ty + 0.29, w: 0.18, h: -0.18, line: { color: DARK, width: 1, endArrowType: 'triangle' } });
		});
	title(s, 'Build a Life on Your Architecture Project', 7.931, 1.387, 4.081, 1.515);
	para(s, LOREM_L, 7.931, 3.189, 4.657, 1.078);
	logo(s, 0.761, 0.329, 0.396);
	t(s, 'Arazuna', { x: 0.319, y: 0.682, w: 1.28, h: 0.337, fontFace: MONT_XB, fontSize: 14, bold: true, color: DARK, align: 'center' });
});

// 9 — creative solutions, two floating cards
slides.push((s) => {
	photo(s, 6.508, 0, 3.413, 5.365);
	photo(s, 10.04, 2.135, 3.294, 5.365);
	title(s, 'Creative Solutions\nProfessionally Designed Architecture', 0.776, 1.72, 4.955, 1.414, { fontSize: 26 });
	para(s, LOREM_L, 0.776, 3.68, 4.955, 1.078);
	para(s, LOREM_M, 0.776, 4.878, 4.955, 0.825);
	logo(s, 0.387, 0.334, 0.396);
	[{ x: 7.069, y: 5.703, bg: RED, fg: WHITE, body: 'F2F2F2', label: 'Minimalist Architecture' },
	 { x: 9.351, y: 0.333, bg: WHITE, fg: DARK, body: GREY, label: 'Modern Architecture' }].forEach((c) => {
		rect(s, c.x, c.y, 3.482, 1.464, c.bg, { shadow: c.bg === RED ? redShadow() : cardShadow() });
		caption(s, c.label, c.x + 0.297, c.y + 0.154, 2.888, { color: c.fg });
		para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore', c.x + 0.297, c.y + 0.474, 2.888, 0.825, { color: c.body });
	});
});

// 10 — hex badge + numbered list
slides.push((s) => {
	photo(s, 4.466, 0, 4.401, 7.5);
	rect(s, 0, 6.422, 4.466, 1.078, DARK);
	hexBadge(s, 2.234, 2.441, RED, 'building');
	caption(s, 'Your Title Here', 1.279, 3.662, 1.909, { align: 'center' });
	para(s, LOREM_S, 0.794, 3.98, 2.878, 1.078, { align: 'center' });
	title(s, 'Build a Life on Your Architecture Project', 9.196, 1.135, 3.873, 1.515);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip', 9.196, 3.085, 3.873, 1.078);
	[['01. Your title here', 4.405, 4.691], ['02. Your title here', 5.506, 5.792]].forEach(([label, ly, by]) => {
		t(s, label, { x: 9.196, y: ly, w: 1.747, h: 0.286, fontFace: POPPINS, fontSize: 11, color: DARK });
		para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt', 9.196, by, 3.206, 0.573);
	});
	t(s, 'www.arazuna.com', { x: 0.792, y: 6.81, w: 2.881, h: 0.303, fontFace: MONT, fontSize: 12, color: WHITE, align: 'center' });
	logo(s, 0.387, 0.334, 0.396);
});

// 11 — expert leader grid
slides.push((s) => {
	title(s, 'Our Expert Leader', 4.175, 0.587, 4.984, 0.572, { color: '000000', align: 'center' });
	para(s, LOREM_S, 3.682, 1.461, 5.97, 0.573, { align: 'center' });
	const people = [
		{ x: 1.24, bar: 1.395, bg: RED, name: 'John Mahrez', role: 'CEO / Founder', nx: 1.587 },
		{ x: 4.035, bar: 4.19, bg: DARK, name: 'Bill William', role: 'Head of Oprational', nx: 4.516 },
		{ x: 6.83, bar: 6.985, bg: DARK, name: 'Allan Senderson', role: 'Head of Development', nx: 7.311 },
		{ x: 9.625, bar: 9.779, bg: DARK, name: 'Liana Hamillton', role: 'Head of Marketing', nx: 10.106 },
	];
	people.forEach((p) => {
		photo(s, p.x, 2.714, 2.469, 3.325);
		rect(s, p.bar, 5.628, 2.158, 0.823, p.bg, { shadow: { type: 'outer', color: p.bg, blur: 37, offset: 24, angle: 90, opacity: 0.15 } });
		t(s, p.name, { x: p.bar, y: 5.719, w: 2.158, h: 0.337, fontFace: MONT, fontSize: 14, bold: true, color: WHITE, align: 'center' });
		t(s, p.role, { x: p.nx, y: 6.05, w: 1.507, h: 0.269, fontSize: 10, italic: true, color: WHITE, align: 'center' });
	});
	logo(s, 0.387, 0.334, 0.396);
});

// 12 — profile with skill bars
slides.push((s) => {
	rect(s, 0, 0, 5.603, 6.643, DARK);
	photo(s, 1.747, 0.857, 4.778, 6.643);
	[['Housing', 1.603], ['Apartment', 3.149], ['Building ', 4.581]].forEach(([label, y]) => {
		t(s, label, { x: 0.135, y, w: 1.437, h: 0.303, rotate: 270, fontFace: MONT, fontSize: 12, bold: true, color: WHITE });
	});
	title(s, 'Allan Senderson', 7.662, 1.395, 4.154, 0.572);
	t(s, 'Professional Architect', { x: 7.662, y: 1.967, w: 2.039, h: 0.303, fontSize: 12, italic: true, color: GREY });
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor', 7.662, 2.513, 4.533, 1.078);
	[['90%', 4.179, 4.556, 3.624], ['80%', 4.987, 5.37, 3.419], ['75%', 5.819, 6.183, 3.0]].forEach(([pct, ly, by, fill]) => {
		caption(s, 'Your Title Here', 7.662, ly, 1.957, { fontSize: 11 });
		t(s, pct, { x: 11.405, y: ly, w: 0.596, h: 0.286, fontFace: MONT, fontSize: 11, bold: true, color: DARK, align: 'right' });
		rect(s, 7.781, by, 4.154, 0.083, SILVER);
		rect(s, 7.781, by, fill, 0.083, RED);
	});
	logo(s, 12.386, 0.327, 0.396);
});

// 13 — expert team cards
slides.push((s) => {
	title(s, 'Our Expert Team', 4.175, 0.451, 4.984, 0.572, { color: '000000', align: 'center' });
	const team = [
		{ x: 1.837, bg: RED, fg: WHITE, body: WHITE, name: 'Alfen Alexander', role: 'Head of Development', nx: 2.134, rx: 2.32 },
		{ x: 5.207, bg: WHITE, fg: DARK, body: GREY, name: 'Noel Sullivan', role: 'Designer', nx: 5.504, rx: 5.689 },
		{ x: 8.576, bg: WHITE, fg: DARK, body: GREY, name: 'Derina Jhonson', role: 'Operations ', nx: 8.873, rx: 9.059 },
	];
	team.forEach((c) => {
		rect(s, c.x, 1.7, 2.92, 4.75, c.bg, { shadow: c.bg === RED ? redShadow() : cardShadow() });
		s.addShape('ellipse', { x: c.x + 0.789, y: 2.112, w: 1.344, h: 1.344, fill: { color: PHOTO } });
		t(s, c.name, { x: c.nx, y: 3.876, w: 2.326, h: 0.337, fontFace: MONT, fontSize: 14, bold: true, color: c.fg, align: 'center' });
		t(s, c.role, { x: c.rx, y: 4.212, w: 1.955, h: 0.269, fontSize: 10, color: c.body === WHITE ? WHITE : GREY, align: 'center' });
		para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt', c.x + 0.175, 4.671, 2.57, 0.825, { align: 'center', color: c.body });
		['twitter', 'facebook', 'instagram'].forEach((icon, i) => {
			const bx = c.x + 1.028 + i * 0.313;
			s.addShape('ellipse', { x: bx, y: 5.798, w: 0.24, h: 0.24, fill: { color: BLUE } });
			glyph(s, icon, bx + 0.12, 5.918, 0.12, WHITE);
		});
	});
	logo(s, 12.386, 0.327, 0.396);
});

// 14 — service grid (2 x 3 icons)
slides.push((s) => {
	title(s, 'Arazuna Service', 4.175, 0.493, 4.984, 0.572, { align: 'center' });
	s.addShape('line', { x: 6.342, y: 1.367, w: 0.65, h: 0, line: { color: RED, width: 2.2 } });
	const icons = ['plan', 'ruler', 'helmet', 'tower', 'building', 'house'];
	icons.forEach((icon, i) => {
		const col = i % 3, row = Math.floor(i / 3);
		const cx = 2.858 + col * 3.808, y = 2.531 + row * 2.628;
		glyph(s, icon, cx, y, 0.55, RED);
		caption(s, 'Your Title Here', cx - 0.95, y + 0.524, 1.909, { align: 'center' });
		para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore. ', cx - 1.511, y + 0.924, 3.031, 0.825, { align: 'center' });
	});
	logo(s, 12.386, 0.327, 0.396);
});

// 15 — four feature cards over a red panel
slides.push((s) => {
	photo(s, 8.687, 0, 4.646, 7.5);
	rect(s, 8.677, 0, 4.657, 7.5, RED, { fill: { color: RED, transparency: 15 } });
	title(s, 'Creative Solutions\nProfessionally Designed Architecture', 1.735, 0.957, 6.177, 1.414, { fontSize: 26 });
	[{ x: 1.735, y: 3.394, icon: 'plan' }, { x: 6.904, y: 3.394, icon: 'ruler' },
	 { x: 1.735, y: 5.29, icon: 'tower' }, { x: 6.904, y: 5.29, icon: 'building' }].forEach((c) => {
		rect(s, c.x, c.y, 4.694, 1.464, WHITE, { shadow: cardShadow() });
		glyph(s, c.icon, c.x + 0.575, c.y + 0.545, 0.55, RED);
		caption(s, 'Your Title Here', c.x + 1.066, c.y + 0.15, 1.909);
		para(s, LOREM_CARD, c.x + 1.068, c.y + 0.497, 3.354, 0.825);
	});
	logo(s, 0.387, 0.334, 0.396);
});

// 16 — best service, hexagon list
slides.push((s) => {
	rect(s, 0, 0, 1.021, 7.5, DARK);
	photo(s, 1.021, 0, 5.16, 7.5);
	title(s, 'Arazuna Best Architecture Service', 7.253, 1.193, 5.06, 1.043);
	[{ y: 3.225, color: RED, icon: 'tower', label: '01. Your Title Here', ty: 2.829, by: 3.159 },
	 { y: 4.509, color: RED, icon: 'house', label: '02. Your Title Here', ty: 4.097, by: 4.4 },
	 { y: 5.777, color: DARK, icon: 'building', label: '03. Your Title Here', ty: 5.375, by: 5.691 }].forEach((r) => {
		s.addShape('hexagon', { x: 7.326, y: r.y - 0.338, w: 0.755, h: 0.676, rotate: 90, fill: { color: WHITE }, shadow: softShadow('000000') });
		s.addShape('hexagon', { x: 7.387, y: r.y - 0.284, w: 0.633, h: 0.567, rotate: 90, fill: { color: r.color } });
		glyph(s, r.icon, 7.703, r.y, 0.29, WHITE);
		caption(s, r.label, 8.201, r.ty, 3.973);
		t(s, 'Lorem ipsum dolor sit amet, consecur adipiscing elit, sed do eiod tempor incididnt ut laborens', { x: 8.201, y: r.by, w: 3.973, h: 0.671, fontSize: 12, lineSpacingMultiple: 1.5 });
	});
	s.addShape('ellipse', { x: 0.344, y: 6.819, w: 0.334, h: 0.334, fill: { type: 'none' }, line: { color: WHITE, width: 1 } });
	s.addShape('line', { x: 0.42, y: 6.94, w: 0.18, h: 0.1, line: { color: WHITE, width: 1 } });
	s.addShape('line', { x: 0.6, y: 7.04, w: -0.18, h: 0.1, line: { color: WHITE, width: 1 } });
	t(s, 'www.arazuna.com', { x: -0.579, y: 1.315, w: 2.178, h: 0.303, rotate: 270, fontFace: MONT, fontSize: 12, color: WHITE, align: 'center' });
	logo(s, 12.386, 0.327, 0.396);
});

// 17 — three numbered strips under a white card
slides.push((s) => {
	rect(s, 0.881, 0.628, 4.802, 4.394, WHITE, { shadow: cardShadow() });
	photo(s, 5.681, 0.628, 6.774, 4.394);
	title(s, 'Build a Life on Your Architecture Project', 1.343, 1.31, 3.873, 1.515, { align: 'center' });
	para(s, LOREM_L, 1.113, 3.211, 4.333, 1.078, { align: 'center' });
	[{ x: 0.879, bg: RED, num: '1', numColor: RED, dx: 1.096, tx: 1.922 },
	 { x: 4.884, bg: DARK, num: '2', numColor: DARK, dx: 5.155, tx: 6.042 },
	 { x: 8.873, bg: DARK, num: '3', numColor: DARK, dx: 9.093, tx: 9.904 }].forEach((c) => {
		rect(s, c.x, 5.361, 3.582, 1.511, c.bg, { shadow: c.bg === RED ? redShadow() : cardShadow() });
		t(s, c.num, { shape: 'ellipse', x: c.dx, y: 5.575, w: 0.614, h: 0.622, fill: { color: WHITE }, fontFace: MONT, fontSize: 16, bold: true, color: c.numColor, align: 'center', valign: 'middle', margin: 0 });
		caption(s, 'Your Title Here', c.tx, 5.58, 1.909, { color: WHITE });
		para(s, LOREM_XS, c.tx, 5.928, 2.333, 0.825, { color: WHITE });
	});
});

// 18 — we give ideas
slides.push((s) => {
	[[1.374, 1.529, 'helmet'], [3.063, 3.182, 'tower'], [4.753, 4.914, 'building']].forEach(([cardY, rowY, icon]) => {
		rect(s, 0.999, cardY, 5.464, 1.458, WHITE, { shadow: cardShadow() });
		rect(s, 1.214, rowY, 1.362, 1.135, RED);
		glyph(s, icon, 1.895, rowY + 0.568, 0.6, WHITE);
		caption(s, 'Your Title Here', 2.955, rowY, 1.909);
		para(s, LOREM_CARD, 2.955, rowY + 0.31, 3.127, 0.825);
	});
	title(s, 'We Give Ideas For Your Dreams', 7.655, 1.501, 4.679, 1.043);
	para(s, LOREM_L, 7.655, 2.995, 4.679, 1.078);
	para(s, LOREM_M, 7.655, 4.194, 4.679, 0.825);
	pillButton(s, 'Learn More  ', 7.762, 5.531);
	logo(s, 12.386, 0.327, 0.396);
});

// 19 — break slide
slides.push((s) => {
	s.background = { color: DARK };
	photo(s, 0, 0.79, 10.185, 4.549);
	rect(s, 6.296, 4.481, 7.037, 2.046, RED, { shadow: redShadow() });
	t(s, QUOTE, { x: 1.064, y: 5.902, w: 3.905, h: 0.625, fontFace: MONT, fontSize: 11, italic: true, color: WHITE, lineSpacingMultiple: 1.5 });
	t(s, 'Break Slide', { x: 7.07, y: 4.949, w: 5.489, h: 1.111, fontFace: MONT_BLACK, fontSize: 60, bold: true, color: WHITE, align: 'center' });
	logo(s, 11.429, 1.802, 0.661);
	t(s, 'Arazuna', { x: 10.775, y: 2.473, w: 1.968, h: 0.539, fontFace: MONT_XB, fontSize: 26, bold: true, color: WHITE, align: 'center' });
});

// 20 — four category cards
slides.push((s) => {
	s.background = { color: RED };
	rect(s, 0, 0, 13.333, 5.416, DARK);
	[['Building', 0.984, 0.984, 1.348, 1.1], ['Apartment', 3.878, 3.883, 4.247, 4.0],
	 ['Housing', 6.772, 6.772, 7.137, 6.89], ['Real Estate', 9.665, 9.665, 10.035, 9.788]]
		.forEach(([label, px, cardX, capX, bodyX]) => {
			photo(s, px, 0.73, 2.684, 4.691);
			rect(s, cardX, 5.421, 2.687, 1.421, WHITE, { shadow: cardShadow() });
			t(s, label, { x: capX, y: 5.561, w: 1.956, h: 0.415, fontFace: MONT, fontSize: 14, bold: true, color: DARK, align: 'center', lineSpacingMultiple: 1.5 });
			para(s, 'Lorem ipsum dolor sit amet, consectetur', bodyX, 6.062, 2.451, 0.573, { align: 'center' });
		});
});

// 21 — dark project panel + three photos
slides.push((s) => {
	rect(s, 0, 0, 7.5, 7.5, DARK);
	[[4.583, PHOTO], [7.5, PHOTO_D], [10.417, PHOTO]].forEach(([x, tone]) => photo(s, x, 2.484, 2.917, 4.337, { fill: { color: tone } }));
	title(s, 'Creative Solutions\nProfessionally Designed Architecture', 0.698, 0.669, 6.177, 1.414, { fontSize: 26, color: WHITE });
	caption(s, 'Our Project ', 0.698, 3.865, 1.909, { color: WHITE });
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. Maecenas accumsan lacus vel facilisis', 0.698, 4.2, 3.203, 1.33, { color: WHITE });
	[[4.583, RED], [7.5, DARK], [10.417, RED]].forEach(([x, color]) => {
		t(s, 'Your Text Here', { shape: 'rect', x, y: 6.325, w: 2.917, h: 0.496, fill: { color }, fontFace: MONT, fontSize: 12, bold: true, color: WHITE, align: 'center', valign: 'middle', margin: 0 });
	});
	logo(s, 12.386, 0.327, 0.396);
});

// 22 — modern property
slides.push((s) => {
	photo(s, 7.431, 0.941, 4.764, 2.882);
	rect(s, 7.431, 3.823, 4.764, 2.736, RED, { shadow: cardShadow() });
	title(s, 'Arazuna Best Modern Property ', 1.139, 0.974, 4.516, 1.043);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. Maecenas accumsan lacus vel', 1.139, 2.377, 5.296, 0.825);
	caption(s, 'Our Project ', 7.714, 4.113, 1.909, { color: WHITE });
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. Maecenas accumsan lacus vel facilisis', 7.714, 4.418, 4.197, 1.078, { color: WHITE });
	s.addShape('roundRect', { x: 9.891, y: 5.776, w: 1.836, h: 0.502, rectRadius: 0.084, fill: { color: WHITE } });
	t(s, '$ 1.745.000', { x: 10.007, y: 5.825, w: 1.605, h: 0.404, fontFace: MONT, fontSize: 18, color: DARK, align: 'center' });
	photo(s, 1.139, 3.823, 2.604, 2.736);
	photo(s, 4.063, 3.823, 2.604, 2.736);
	logo(s, 0.387, 0.334, 0.396);
});

// 23 — popular property grid
slides.push((s) => {
	photo(s, 0, 0, 4.28, 7.5);
	rect(s, 0, 0, 4.28, 7.5, RED, { fill: { color: RED, transparency: 15 } });
	title(s, 'Arazuna Popular Property ', 0.284, 1.946, 3.624, 1.043, { color: WHITE });
	para(s, LOREM_L, 0.284, 3.406, 3.624, 1.33, { color: WHITE });
	caption(s, 'Learn More  ', 0.284, 5.153, 1.217, { color: WHITE });
	arrow(s, 1.581, 5.304, 0.549, WHITE);
	[[5.067, 0.739, 2.553], [9.016, 0.734, 2.553], [5.067, 3.935, 5.787], [9.016, 3.935, 5.787]].forEach(([x, y, cardY]) => {
		photo(s, x, y, 3.568, 1.819);
		rect(s, x, cardY, 3.568, 1.012, WHITE, { shadow: cardShadow() });
		caption(s, 'Your Title Here', x + 0.24, y + 2.017, 1.909);
		t(s, 'Lorem ipsum dolor sit amet, consectetur', { x: x + 0.24, y: y + 2.326, w: 3.089, h: 0.32, lineSpacingMultiple: 1.5 });
	});
});

// 24 — laptop mockup
slides.push((s) => {
	rect(s, 9.653, 0, 3.681, 7.5, SILVER);
	title(s, 'Architecture is a way of thinking about the world', 0.865, 1.803, 4.509, 1.414, { fontSize: 26 });
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. Maecenas accumsan lacus vel', 0.906, 3.379, 4.739, 0.825);
	[['1170+', 0.905], ['90%', 3.465]].forEach(([big, x]) => {
		t(s, big, { x, y: 4.525, w: 1.348, h: 0.404, fontFace: MONT, fontSize: 18, bold: true, color: RED });
		caption(s, 'Your Title Here', x, 4.95, 1.909);
		para(s, LOREM_TINY, x, 5.298, 2.178, 0.573);
	});
	// laptop mockup: rounded dark lid, grey screen, silver base with a trackpad notch
	s.addShape('roundRect', { x: 7.14, y: 1.167, w: 7.11, h: 4.8, rectRadius: 0.13, fill: { color: '0A0A0C' } });
	photo(s, 7.226, 1.365, 6.937, 4.413);
	s.addShape('roundRect', { x: 6.43, y: 6.16, w: 8.6, h: 0.15, rectRadius: 0.07, fill: { color: '86888C' } });
	rect(s, 9.6, 6.16, 1.6, 0.08, '6E7074');
	logo(s, 0.387, 0.334, 0.396);
});

// 25 — phone mockup
slides.push((s) => {
	rect(s, 0, 0, 3.361, 7.5, RED);
	// phone: dark shell with rounded corners + grey screen
	s.addShape('roundRect', { x: 1.47, y: 0.641, w: 3.766, h: 7.552, rectRadius: 0.45, fill: { color: '111111' } });
	s.addShape('roundRect', { x: 1.694, y: 0.764, w: 3.361, h: 7.222, rectRadius: 0.34, fill: { color: PHOTO } });
	title(s, 'Architecture is a way of thinking about the world', 6.778, 1.817, 5.177, 1.515);
	para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. Maecenas accumsan lacus vel facilisis volutpat est velit egestas dui. Lorem ipsum dolor sit amet ', 6.778, 3.75, 5.177, 1.078);
	[['1170+', 'Coustomer', 6.378, 6.659, 1.348], ['12K', 'Project Success', 8.262, 8.711, 1.011], ['1115', 'Happy Client ', 10.158, 10.594, 1.011]]
		.forEach(([big, small, gx, bx, bw]) => {
			t(s, big, { x: bx, y: 5.246, w: bw, h: 0.505, fontFace: MONT, fontSize: 24, bold: true, color: RED, align: 'center' });
			t(s, small, { x: gx, y: 5.779, w: 1.909, h: 0.286, fontFace: MONT, fontSize: 11, bold: true, color: DARK, align: 'center' });
		});
	[8.25, 10.133].forEach((x) => s.addShape('roundRect', { x, y: 5.339, w: 0.05, h: 0.632, rectRadius: 0.025, fill: { color: SILVER } }));
	logo(s, 12.386, 0.327, 0.396);
});

// 26 — desktop mockup with progress rings
slides.push((s) => {
	photo(s, 0, 0, 13.333, 4.955, { fill: { color: PHOTO_D } });
	rect(s, 0, 0, 13.333, 4.955, DARK, { fill: { color: DARK, transparency: 15 } });
	title(s, 'Move Forward with Our Agency', 4.012, 0.584, 5.309, 1.043, { color: WHITE, align: 'center' });
	// monitor mockup: dark bezel, grey screen, silver chin, tapered neck and foot
	s.addShape('roundRect', { x: 3.79, y: 2.28, w: 5.75, h: 3.42, rectRadius: 0.06, fill: { color: '070707' } });
	photo(s, 3.972, 2.458, 5.403, 3.014);
	rect(s, 3.79, 5.7, 5.75, 0.45, 'D3D3D3');
	s.addShape('trapezoid', { x: 6.2, y: 6.15, w: 0.93, h: 0.63, fill: { color: 'A9A9A9' } });
	s.addShape('ellipse', { x: 5.6, y: 6.7, w: 2.13, h: 0.3, fill: { color: 'C4C4C4' } });
	[[1.13, 'Your Title Here', '85%', 0.748, 0.768], [10.85, 'Your Title Here', '15k', 10.487, 10.489]].forEach(([rx, label, value, capX, bodyX]) => {
		s.addShape('ellipse', { x: rx + 0.07, y: 3.82, w: 0.715, h: 0.715, fill: { color: WHITE, transparency: 15 } });
		s.addShape('ellipse', { x: rx, y: 3.75, w: 0.854, h: 0.854, fill: { type: 'none' }, line: { color: WHITE, width: 3 } });
		s.addShape('arc', { x: rx, y: 3.75, w: 0.854, h: 0.854, angleRange: [344, 268], fill: { type: 'none' }, line: { color: RED, width: 3 } });
		t(s, value, { x: rx + 0.07, y: 4.009, w: 0.715, h: 0.337, fontFace: MONT, fontSize: 14, bold: true, color: '000000', align: 'center' });
		caption(s, label, capX, 5.26, 1.909);
		para(s, LOREM_TINY, bodyX, 5.59, 2.431, 0.573);
	});
});

// 27 — testimonials
slides.push((s) => {
	title(s, 'Our Customer Says', 4.012, 0.408, 5.309, 0.572, { align: 'center' });
	para(s, LOREM_S, 3.319, 1.309, 6.695, 0.573, { align: 'center' });
	[{ x: 0.892, bg: WHITE, fg: DARK, body: GREY, name: 'Allen Christine', nx: 2.031, qx: 4.144 },
	 { x: 4.607, bg: RED, fg: WHITE, body: WHITE, name: 'Adam Amstrong', nx: 5.657, qx: 7.858 },
	 { x: 8.321, bg: WHITE, fg: DARK, body: GREY, name: 'Jody Lincon', nx: 9.465, qx: 11.572 }].forEach((c) => {
		rect(s, c.x + 0.377, 2.523, 3.365, 4.12, c.bg, { shadow: c.bg === RED ? redShadow() : cardShadow() });
		s.addShape('ellipse', { x: c.x + 1.402, y: 2.853, w: 1.317, h: 1.317, fill: { color: PHOTO } });
		t(s, c.name, { x: c.nx, y: 4.518, w: c.bg === RED ? 2.019 : 1.843, h: 0.337, fontFace: MONT, fontSize: 14, bold: true, color: c.fg, align: 'center' });
		para(s, 'Nullam ac tortor vitae purus faucibus ornare suspendisse sed nisi. Maecenas accumsan lacus vel facilisis volutpat', c.x + 0.728, 5.038, 2.665, 1.175, { fontSize: 11, align: 'center', color: c.body });
		[0, 0.169].forEach((dx) => s.addShape('custGeom', {
			x: c.qx + dx, y: 6.505, w: 0.113, h: 0.277, fill: { color: c.bg === RED ? WHITE : RED },
			points: [{ x: 0, y: 0 }, { x: 0.113, y: 0 }, { x: 0.113, y: 0.15 }, { x: 0.05, y: 0.277 }, { x: 0, y: 0.277 }, { close: true }],
		}));
	});
	logo(s, 12.386, 0.327, 0.396);
});

// 28 — infographic: chevron arrows with icon discs
slides.push((s) => {
	title(s, 'Our Infographic', 4.099, 0.922, 5.309, 0.64, { fontSize: 32, align: 'center' });
	[{ x: 1.478, color: RED, shade: 'DA0000', icon: 'plan', capX: 1.652, bodyX: 1.25, bodyW: 2.58, label: '01. Your Title Here', body: LOREM_XS },
	 { x: 4.188, color: DARK, shade: '000000', icon: 'helmet', capX: 4.36, bodyX: 4.099, bodyW: 2.447, label: '02. Your Title Here', body: LOREM_XS + ' inci' },
	 { x: 6.892, color: RED, shade: 'DA0000', icon: 'tower', capX: 7.067, bodyX: 6.807, bodyW: 2.447, label: '03. Your Title Here', body: LOREM_XS },
	 { x: 9.597, color: DARK, shade: '000000', icon: 'building', capX: 9.774, bodyX: 9.497, bodyW: 2.447, label: '04. Your Title Here', body: LOREM_XS }]
		.forEach((c) => {
			s.addShape('chevron', { x: c.x, y: 4.095, w: 2.259, h: 0.609, fill: { color: 'F5F6F8' } });   // drop shadow band
			s.addShape('chevron', { x: c.x, y: 3.911, w: 2.259, h: 0.604, fill: { color: c.color } });
			rect(s, c.x, 4.515, 1.955, 0.095, c.shade);                                                   // bevel under arrow
			s.addShape('ellipse', { x: c.x + 0.476, y: 4.067, w: 1.306, h: 0.292, fill: { color: c.shade } }); // disc shadow
			s.addShape('ellipse', { x: c.x + 0.476, y: 2.907, w: 1.306, h: 1.307, fill: { color: c.color }, shadow: dropShadow() });
			glyph(s, c.icon, c.x + 1.129, 3.561, 0.543, WHITE);
			caption(s, c.label, c.capX, 5.311, 1.909, { align: 'center' });
			para(s, c.body, c.bodyX, 5.614, c.bodyW, 0.825, { align: 'center' });
		});
	logo(s, 12.386, 0.327, 0.396);
});

// 29 — infographic: segmented percentage rings
slides.push((s) => {
	title(s, 'Our Infographic', 4.099, 0.922, 5.309, 0.64, { fontSize: 32, align: 'center' });
	[{ cx: 2.213, pct: 60, color: RED, dim: 'FBD3D3', icon: 'house', txt: '60%', tx: 1.682, tw: 1.063, capX: 1.272, bodyX: 0.914 },
	 { cx: 5.155, pct: 50, color: DARK, dim: 'D8D9DB', icon: 'star', txt: '50%', tx: 4.633, tw: 1.045, capX: 4.192, bodyX: 3.879 },
	 { cx: 8.174, pct: 90, color: RED, dim: 'FBD3D3', icon: 'drop', txt: '90%', tx: 7.643, tw: 1.063, capX: 7.233, bodyX: 6.875 },
	 { cx: 11.117, pct: 40, color: DARK, dim: 'D8D9DB', icon: 'database', txt: '40%', tx: 10.576, tw: 1.082, capX: 10.153, bodyX: 9.813 }]
		.forEach((r) => {
			const cy = 3.471, R = 0.79, lit = Math.round((r.pct / 100) * 16);
			for (let i = 0; i < 16; i++) {                                     // 16 tick segments around the ring
				const ang = i * 22.5, rad = (ang * Math.PI) / 180;
				s.addShape('rect', {
					x: r.cx + R * Math.sin(rad) - 0.115, y: cy - R * Math.cos(rad) - 0.075,
					w: 0.23, h: 0.15, rotate: ang, fill: { color: i < lit ? r.color : r.dim },
				});
			}
			s.addShape('ellipse', { x: r.cx - 0.655, y: cy - 0.655, w: 1.311, h: 1.311, fill: { color: r.color } });
			glyph(s, r.icon, r.cx, cy, 0.42, WHITE);
			t(s, r.txt, { x: r.tx, y: 4.765, w: r.tw, h: 0.572, fontFace: MONT, fontSize: 28, bold: true, color: DARK, align: 'center' });
			caption(s, '01. Your Title Here', r.capX, 5.647, 1.909, { align: 'center' });
			para(s, LOREM_XS, r.bodyX, 5.94, 2.58, 0.825, { align: 'center' });
		});
	logo(s, 12.386, 0.327, 0.396);
});

// 30 — infographic: pyramid
slides.push((s) => {
	title(s, 'Our Infographic', 4.099, 0.922, 5.309, 0.64, { fontSize: 32, align: 'center' });
	const tiers = [
		{ y: 2.43, w: 1.274, h: 1.031, color: DARK, icon: 'send', shape: 'triangle' },
		{ y: 3.838, w: 2.384, h: 0.649, color: RED, icon: 'clock', shape: 'trapezoid' },
		{ y: 4.878, w: 3.499, h: 0.649, color: DARK, icon: 'gear', shape: 'trapezoid' },
		{ y: 5.919, w: 4.614, h: 0.649, color: RED, icon: 'house', shape: 'trapezoid' },
	];
	tiers.forEach((tr) => {                                                     // grey offset copy = drop shadow
		s.addShape(tr.shape, { x: 6.721 - tr.w / 2 - 0.2, y: tr.y + 0.12, w: tr.w, h: tr.h, fill: { color: SILVER } });
		s.addShape(tr.shape, { x: 6.721 - tr.w / 2, y: tr.y, w: tr.w, h: tr.h, fill: { color: tr.color } });
		glyph(s, tr.icon, 6.721, tr.y + tr.h - 0.31, 0.34, WHITE);
	});
	[{ n: '1', color: DARK, x: 1.28, y: 2.736, tx: 1.18, ty: 3.337, label: 'Your Step One', bw: 1.894 },
	 { n: '2', color: RED, x: 10.009, y: 2.742, tx: 9.907, ty: 3.337, label: 'Your Step Two', bw: 2.424 },
	 { n: '3', color: DARK, x: 1.28, y: 4.878, tx: 1.18, ty: 5.477, label: 'Your Step Three', bw: 1.894 },
	 { n: '4', color: RED, x: 10.009, y: 4.884, tx: 9.906, ty: 5.483, label: 'Your Step Four', bw: 2.424 }]
		.forEach((st) => {
			t(s, st.n, { shape: 'ellipse', x: st.x, y: st.y, w: 0.458, h: 0.458, fill: { color: st.color }, fontFace: MONT, fontSize: 16, bold: true, color: WHITE, align: 'center', valign: 'middle', margin: 0 });
			t(s, st.label, { x: st.tx, y: st.ty, w: 1.85, h: 0.337, fontFace: MONT, fontSize: 14, color: DARK });
			t(s, 'Lorem ipsum dolor, consect \nadipiscing elit roin cursu.', { x: st.tx, y: st.ty + 0.335, w: st.bw, h: 0.546, lineSpacingMultiple: 1.4 });
		});
	logo(s, 12.386, 0.327, 0.396);
});

// 31 — infographic: four pie wedges
slides.push((s) => {
	title(s, 'Our Infographic', 4.099, 0.922, 5.309, 0.64, { fontSize: 32, align: 'center' });
	const cx = 6.662, cy = 4.429, R = 1.61;
	[{ range: [225, 315], color: RED, icon: 'gear', dx: 0, dy: -1 },
	 { range: [315, 45], color: DARK, icon: 'chart', dx: 1, dy: 0 },
	 { range: [45, 135], color: RED, icon: 'bulb', dx: 0, dy: 1 },
	 { range: [135, 225], color: DARK, icon: 'note', dx: -1, dy: 0 }].forEach((w) => {
		const g = 0.045;                                                       // radial gap between wedges
		s.addShape('pie', { x: cx - R + w.dx * g, y: cy - R + w.dy * g, w: 2 * R, h: 2 * R, angleRange: w.range, fill: { color: w.color } });
		glyph(s, w.icon, cx + w.dx * 0.87, cy + w.dy * 0.87, 0.42, WHITE);
	});
	[{ v: '145', vx: 3.44, vw: 0.684, lx: 2.304, ly: 3.079, bx: 1.4, by: 3.417, align: 'right' },
	 { v: '1.200', vx: 3.194, vw: 0.954, lx: 2.328, ly: 4.955, bx: 1.4, by: 5.291, align: 'right' },
	 { v: '1.2 M', vx: 9.175, vw: 0.919, lx: 9.175, ly: 3.08, bx: 9.166, by: 3.417, align: 'left' },
	 { v: '80%', vx: 9.166, vw: 0.837, lx: 9.175, ly: 4.96, bx: 9.168, by: 5.291, align: 'left' }].forEach((m) => {
		t(s, m.v, { x: m.vx, y: m.ly - 0.404, w: m.vw, h: 0.438, fontFace: MONT, fontSize: 20, bold: true, color: RED, align: m.align });
		t(s, 'Write your title', { x: m.lx, y: m.ly, w: 1.844, h: 0.337, fontFace: MONT, fontSize: 14, bold: true, color: DARK, align: m.align });
		para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do', m.bx, m.by, 2.765, 0.573, { align: m.align });
	});
	logo(s, 12.386, 0.327, 0.396);
});

// 32 — infographic: four interlocking arrows
slides.push((s) => {
	title(s, 'Our Infographic', 4.099, 0.922, 5.309, 0.64, { fontSize: 32, align: 'center' });
	// four overlapping C-rings around a common centre
	const cx = 6.667, cy = 4.517, RING = 2.35, OFF = 0.85;
	[{ dx: -1, dy: -1, range: [90, 0], color: RED },
	 { dx: 1, dy: -1, range: [180, 90], color: DARK },
	 { dx: 1, dy: 1, range: [270, 180], color: RED },
	 { dx: -1, dy: 1, range: [0, 270], color: DARK }].forEach((q) => {
		s.addShape('blockArc', {
			x: cx + q.dx * OFF - RING / 2, y: cy + q.dy * OFF - RING / 2, w: RING, h: RING,
			angleRange: q.range, arcThicknessRatio: 0.55, fill: { color: q.color },
		});
	});
	// flat arrow heads converging on the centre, leaving a white cross between the tips.
	// Rotation happens about the shape centre, so each head is placed by its centre point.
	const HW = 1.167, HH = 0.587, GAP = 0.2, MID = GAP + HH / 2;
	[{ rot: 180, color: RED, mx: 0, my: -MID },
	 { rot: 0, color: RED, mx: 0, my: MID },
	 { rot: 90, color: DARK, mx: -MID, my: 0 },
	 { rot: 270, color: DARK, mx: MID, my: 0 }].forEach((h) => {
		s.addShape('triangle', { x: cx + h.mx - HW / 2, y: cy + h.my - HH / 2, w: HW, h: HH, rotate: h.rot, fill: { color: h.color } });
	});
	[{ x: 0.86, y: 2.954, color: RED, icon: 'gear', tx: 1.979, ty: 3.021, bx: 1.979, by: 3.358 },
	 { x: 0.86, y: 5.023, color: DARK, icon: 'target', tx: 1.995, ty: 5.124, bx: 1.986, by: 5.427 },
	 { x: 9.23, y: 2.954, color: DARK, icon: 'handshake', tx: 10.365, ty: 3.021, bx: 10.37, by: 3.358 },
	 { x: 9.23, y: 5.023, color: RED, icon: 'growth', tx: 10.365, ty: 5.033, bx: 10.37, by: 5.427 }].forEach((m) => {
		s.addShape('triangle', { x: m.x, y: m.y, w: 0.911, h: 0.943, rotate: 90, fill: { color: m.color } });
		glyph(s, m.icon, m.x + 0.455, m.y + 0.471, 0.4, WHITE);
		caption(s, 'Your Title Here', m.tx, m.ty, 2.139);
		t(s, 'Vestibulum neque elit, Class aptent taciti sociosqu ad litora torquent per conubia', { x: m.bx, y: m.by, w: 2.286, h: 0.686, lineSpacingMultiple: 1.5 });
	});
	logo(s, 12.386, 0.327, 0.396);
});

// 33 — contact
slides.push((s) => {
	rect(s, 7.323, 1.066, 6.01, 6.434, DARK);
	photo(s, 7.323, 0, 6.01, 6.434);
	rect(s, 0, 1.066, 7.323, 5.368, RED);
	title(s, 'Let\u2019s Get in Touch With Arazuna Architecture', 1.288, 1.801, 4.746, 0.976, { fontSize: 26, color: WHITE });
	para(s, LOREM_M, 1.288, 3.168, 4.746, 0.825, { color: WHITE });
	[['Our Address', '13th Street, New York, NY 10011, USA.', 4.186, 4.51, 3.956],
	 ['Phone & Email', '+62 205 7516  / arazuna@gmail.com', 5.011, 5.336, 2.895]].forEach(([head, body, hy, by, bw]) => {
		t(s, head, { x: 1.288, y: hy, w: 2.125, h: 0.324, fontFace: MONT, fontSize: 12, bold: true, color: WHITE, lineSpacingMultiple: 1.2 });
		t(s, body, { x: 1.288, y: by, w: bw, h: 0.348, fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
	});
	t(s, 'www.arazuna.com', { x: 8.887, y: 6.816, w: 2.881, h: 0.303, fontFace: MONT, fontSize: 12, color: WHITE, align: 'center' });
	logo(s, 0.387, 0.334, 0.396);
});

// 34 — thank you
slides.push((s) => {
	photo(s, 0, 0, 13.333, 7.5);
	rect(s, 0, 0, 13.333, 7.5, RED, { fill: { color: RED, transparency: 15 } });
	s.addShape('rect', { x: 2.097, y: 2.569, w: 9.139, h: 2.361, fill: { type: 'none' }, line: { color: WHITE, width: 1.5 } });
	t(s, 'Thank You', { x: 2.989, y: 3.195, w: 7.355, h: 1.111, fontFace: MONT_BLACK, fontSize: 60, bold: true, color: WHITE, align: 'center' });
});

// ------------------------------------------------------------------ build --
const pptx = new pptxgen();
pptx.defineLayout({ name: 'ARAZUNA', width: 13.333, height: 7.5 });
pptx.layout = 'ARAZUNA';
pptx.author = 'Arazuna';
pptx.title = 'Arazuna - Architecture & Property Agency Presentation';

slides.forEach((build) => {
	const slide = pptx.addSlide();
	slide.background = { color: WHITE };
	build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '153f1384-dfb0-4724-9c0a-632c6bbc2c7b_grok_final.pptx') });
