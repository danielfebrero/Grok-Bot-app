/**
 * Golf Course — Sport Presentation Template (20 slides, 13.333 x 7.5 in)
 * Rebuilt with pptxgenjs. Raster photos in the source deck are replaced by
 * light-gray "[image]" placeholder rectangles of the same position and size.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');

/* ------------------------------------------------------------------ palette */
const NAVY = '0F112A'; // theme accent1
const GREEN = '02F893'; // theme accent2 - gradient start
const CYAN = '01D4DC'; // theme accent3 - gradient end
const WHITE = 'FFFFFF';
const INK = '262626'; // text1 lum 85%  - body copy
const INK_DK = '0D0D0D'; // text1 lum 95%
const NAVY_75 = '2D337D'; // accent1 lumMod 75 / lumOff 25
const NAVY_60 = '3F47B0'; // accent1 lumMod 60 / lumOff 40
const NAVY_50 = '5860C3'; // accent1 lumMod 50 / lumOff 50
const NAVY_D75 = '0B0D20'; // accent1 lumMod 75
const NAVY_D50 = '070915'; // accent1 lumMod 50
const GREEN_D75 = '01BA6E'; // accent2 lumMod 75
const CYAN_D50 = '006A6E'; // accent3 lumMod 50
const CYAN_L40 = '52F8FE'; // accent3 lumMod 60 / lumOff 40
const GREY_85 = 'D9D9D9'; // background1 lumMod 85
const IMG_BG = 'CCCCCC'; // photo placeholder body
const IMG_FG = '9E9E9E'; // photo placeholder caption

const FONT = 'Work Sans';
const DISPLAY = 'Anton';
const SYM = 'DejaVu Sans'; // pictographic glyphs used for the small icons

/* ------------------------------------------------------------------ helpers */
const NOLINE = { type: 'none' };

/** linear blend of two hex colours, t = 0..1 */
function mix (from, to, t) {
	const chan = i => Math.round(
		parseInt(from.substr(i, 2), 16) * (1 - t) + parseInt(to.substr(i, 2), 16) * t
	).toString(16).padStart(2, '0');
	return (chan(0) + chan(2) + chan(4)).toUpperCase();
}

/** solid rectangle */
function box (s, x, y, w, h, fill, opt) {
	s.addShape('rect', Object.assign({ x, y, w, h, fill: { color: fill }, line: NOLINE }, opt));
}

/**
 * Drop shadow. Always returns a fresh object: pptxgenjs rewrites the shadow
 * values in place while serialising, so a shared literal would be corrupted
 * the second time it is used.
 */
function shadow (blur, offset, opacity) {
	return { type: 'outer', color: '000000', opacity: opacity === undefined ? 0.2 : opacity, blur, offset, angle: 45 };
}

/**
 * The deck's signature 45-degree accent2 -> accent3 gradient.
 * pptxgenjs has no gradient fill, so it is approximated with a vertical colour
 * ramp overlaid by a half-transparent horizontal ramp of the same colours.
 */
function gradient (s, x, y, w, h, steps) {
	const n = steps || 14;
	const bleed = 0.012;
	for (let i = 0; i < n; i++) {
		box(s, x, y + h * i / n, w, h / n + bleed, mix(GREEN, CYAN, (i + 0.5) / n));
	}
	for (let i = 0; i < n; i++) {
		s.addShape('rect', {
			x: x + w * i / n, y, w: w / n + bleed, h,
			fill: { color: mix(GREEN, CYAN, (i + 0.5) / n), transparency: 50 }, line: NOLINE
		});
	}
}

/** text box: top-anchored like PowerPoint, Work Sans by default */
function text (s, content, o) {
	s.addText(content, Object.assign({ fontFace: FONT, valign: 'top', color: INK }, o));
}

/** two-tone display heading, e.g. "Best Our" + "Facilities" */
function heading (s, dark, light, o) {
	text(s, [
		{ text: dark, options: { color: NAVY } },
		{ text: light, options: { color: CYAN } }
	], Object.assign({ fontFace: DISPLAY, fontSize: 37, h: 0.724 }, o));
}

/** replacement for a photo from the original deck */
function photo (s, x, y, w, h) {
	box(s, x, y, w, h, IMG_BG);
	if (w > 0.9 && h > 0.5) {
		text(s, '[image]', {
			x, y: y + h / 2 - 0.18, w, h: 0.36,
			align: 'center', valign: 'middle', fontSize: 12, color: IMG_FG
		});
	}
}

/** small pictographic icon drawn with a unicode glyph */
function glyph (s, ch, x, y, size, color, pt) {
	text(s, ch, {
		x, y, w: size, h: size, align: 'center', valign: 'middle',
		fontFace: SYM, fontSize: pt || Math.round(size * 58), color
	});
}

/** paper-plane / send icon */
function planeIcon (s, cx, cy, r, color) {
	s.addShape('custGeom', {
		x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color }, line: NOLINE,
		points: [{ x: 2 * r, y: 0 }, { x: 0, y: r * 0.95 }, { x: r * 0.78, y: r * 1.2 },
			{ x: r * 0.78, y: 2 * r }, { x: r * 1.22, y: r * 1.45 }, { close: true }]
	});
}

/** globe: circle outline with a meridian and an equator */
function globeIcon (s, cx, cy, r, color) {
	s.addShape('ellipse', { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { type: 'none' }, line: { color, width: 1.5 } });
	s.addShape('ellipse', { x: cx - r * 0.42, y: cy - r, w: r * 0.84, h: 2 * r, fill: { type: 'none' }, line: { color, width: 1.25 } });
	s.addShape('line', { x: cx - r, y: cy, w: 2 * r, h: 0, line: { color, width: 1.25 } });
}

/** house / address icon; `hole` fills the doorway so it reads as a cut-out */
function houseIcon (s, cx, cy, r, color, hole) {
	s.addShape('triangle', { x: cx - r, y: cy - r, w: 2 * r, h: r * 0.95, fill: { color }, line: NOLINE });
	box(s, cx - r * 0.72, cy - r * 0.1, r * 1.44, r * 1.0, color);
	box(s, cx - r * 0.20, cy + r * 0.35, r * 0.40, r * 0.55, hole);
}

/** concentric-ring target used by several infographic slides */
function target (s, cx, cy, r, color, gap) {
	const alt = gap || WHITE;
	[[1, color], [0.72, alt], [0.5, color], [0.26, alt], [0.14, color]].forEach(([k, c]) => {
		s.addShape('ellipse', { x: cx - r * k, y: cy - r * k, w: 2 * r * k, h: 2 * r * k, fill: { color: c }, line: NOLINE });
	});
}

/**
 * Point list for the OOXML `upArrow` preset with explicit adjust values
 * (pptxgenjs cannot pass them through). a1 = shaft width, a2 = head height,
 * both as 1/100000 of the shorter side.
 */
function upArrowPoints (w, h, a1, a2) {
	const ss = Math.min(w, h);
	const half = ss * a1 / 200000;
	const head = ss * a2 / 100000;
	return [{ x: 0, y: head }, { x: w / 2, y: 0 }, { x: w, y: head },
		{ x: w / 2 + half, y: head }, { x: w / 2 + half, y: h },
		{ x: w / 2 - half, y: h }, { x: w / 2 - half, y: head }, { close: true }];
}

/** Point list for the OOXML `parallelogram` preset with an explicit adjust value */
function parallelogramPoints (w, h, adj) {
	const dx = Math.min(w, h) * adj / 100000;
	return [{ x: 0, y: h }, { x: dx, y: 0 }, { x: w, y: 0 }, { x: w - dx, y: h }, { close: true }];
}

/** ring-bound notebook pictogram */
function notebook (s, cx, cy, w, color) {
	const h = w * 1.08, l = cx - w / 2, t = cy - h / 2;
	s.addShape('roundRect', { x: l + w * 0.14, y: t, w: w * 0.86, h, rectRadius: 0.02, fill: { color }, line: NOLINE });
	box(s, l + w * 0.36, t + h * 0.14, w * 0.44, h * 0.24, WHITE);
	[0.14, 0.42, 0.70].forEach(k => {
		s.addShape('roundRect', { x: l, y: t + h * k, w: w * 0.30, h: h * 0.11, rectRadius: 0.01, fill: { color: WHITE }, line: NOLINE });
	});
}

/** dart flying in from the upper left and landing on a target's centre */
function dart (s, cx, cy, len, color) {
	const t = len * 0.13;
	s.addShape('custGeom', {
		x: cx - len, y: cy - len, w: len, h: len, fill: { color }, line: NOLINE,
		points: [
			{ x: 0, y: 0 }, { x: len * 0.42, y: len * 0.10 }, { x: len, y: len * 0.68 },
			{ x: len - t, y: len }, { x: len * 0.10, y: len * 0.42 }, { close: true }
		]
	});
}

/* ------------------------------------------------------- shared top nav bar */
const NAV_LINKS = [
	{ label: 'Profile', x: 7.3498, w: 0.6648 },
	{ label: 'About Us', x: 8.5723, w: 0.8629 },
	{ label: 'Categories', x: 9.9929, w: 0.9593 },
	{ label: 'Contact', x: 11.5099, w: 0.7717 }
];

/**
 * opts.color      - colour of FR / EN and the first links (default INK)
 * opts.linkColor  - override for individual links (array of 4)
 * opts.rule       - colour of the thin rule between FR and EN
 * opts.burger     - hamburger colour: 'gradient' (default) or a hex value
 * opts.burgerX    - hamburger left edge (default 1.1565)
 */
function navBar (s, opts) {
	const o = opts || {};
	const c = o.color || INK;
	const links = o.linkColor || [c, c, c, c];

	// hamburger: three rounded bars
	const bx = o.burgerX === undefined ? 1.1565 : o.burgerX;
	[0, 0.08, 0.16].forEach(dy => {
		if (o.burger === 'flat' || o.burger === undefined) {
			s.addShape('roundRect', { x: bx, y: 0.5075 + dy, w: 0.3227, h: 0.03, rectRadius: 0.015, fill: { color: mix(GREEN, CYAN, 0.5) }, line: NOLINE });
		} else {
			s.addShape('roundRect', { x: bx, y: 0.5075 + dy, w: 0.3227, h: 0.03, rectRadius: 0.015, fill: { color: o.burger }, line: NOLINE });
		}
	});

	text(s, 'FR', { x: 3.4155, y: 0.4448, w: 0.4176, h: 0.3029, fontSize: 12, bold: true, color: c, wrap: false });
	text(s, 'EN', { x: 5.5579, y: 0.4448, w: 0.4334, h: 0.3029, fontSize: 12, bold: true, color: c, wrap: false });
	s.addShape('line', { x: 4.0069, y: 0.5962, w: 1.2986, h: 0, line: { color: o.rule || NAVY, width: 0.75 } });
	gradient(s, 4.2828, 0.5712, 0.747, 0.05, 6);

	NAV_LINKS.forEach((l, i) => {
		text(s, l.label, { x: l.x, y: i === 3 ? 0.4619 : 0.4574, w: l.w, h: 0.2777, fontSize: 10.5, color: links[i], wrap: false });
	});
}

/** body copy shared by nearly every slide */
function body (s, str, x, y, w, h, o) {
	text(s, str, Object.assign({ x, y, w, h, fontSize: 10.5, align: 'justify', lineSpacingMultiple: 1.5 }, o));
}

/* =============================================================== the slides */

const LOREM_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore reprehenderit in voluptate velit esse cillum dolore eu fugiat. consectetur adipiscing elit, sed do eiusmod tempor incididunt';
const LOREM_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt';
const LOREM_SHORT = 'Lorem ipsum dolor sit amet, conse ctetur adipiscing elit, ';
const LOREM_FAC = 'Lorem ipsum dolor sit amet, consectetur adipiscing et dolore incididunt';
const LOREM_VIS = 'Lorem ipsum dolor sit amet, consectetur adipiscing incididunt ut labore et dolore sit amet';
const LOREM_MAEC = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, magna sed pulvinar osuere, magna magna posuere, magna sed';

/* --- 1. cover ------------------------------------------------------------ */
function slide01 (s) {
	box(s, 0, 0, 6.6667, 3.75, NAVY);
	gradient(s, 6.6667, 0, 6.6667, 1.0316);
	photo(s, 3.5368, 1.0316, 9.7965, 5.0947);
	box(s, 10.6105, 1.6211, 1.7789, 2.1289, WHITE); // white frame behind the inset photo

	text(s, [
		{ text: 'GOLF ', options: { color: NAVY } },
		{ text: 'COURSE', options: { color: WHITE } }
	], { x: 1.021, y: 4.8272, w: 6.6445, h: 1.7166, fontFace: DISPLAY, fontSize: 96, wrap: false });

	text(s, 'Sport Presentation Template', { x: 1.0553, y: 0.9741, w: 2.2433, h: 1.0098, fontSize: 18, bold: true, color: WHITE });
	body(s, LOREM_SHORT, 1.0634, 2.3056, 1.9158, 0.8665, { color: WHITE });

	navBar(s, { color: WHITE, rule: WHITE, burger: WHITE, linkColor: [WHITE, WHITE, WHITE, WHITE] });
	photo(s, 10.7716, 1.7912, 1.4569, 1.7887);
}

/* --- 2. welcome ---------------------------------------------------------- */
function slide02 (s) {
	gradient(s, 9.7895, 0, 3.5439, 7.5, 18);
	navBar(s, { linkColor: [INK, INK, WHITE, WHITE] });

	heading(s, 'Welcome To ', 'Golf Course', { x: 1.0604, y: 1.3158, w: 5.6063 });
	text(s, 'About Our Planning', { x: 1.0659, y: 2.6486, w: 3.0447, h: 0.4039, fontSize: 18, bold: true, color: NAVY });
	body(s, LOREM_LONG, 1.0604, 3.3862, 4.9309, 1.1316);

	photo(s, 1.1565, 5.3263, 4.7347, 2.1737);
	photo(s, 7.4421, 1.4, 3.9579, 4.7053);
}

/* --- 3. vision & mission ------------------------------------------------- */
function slide03 (s) {
	photo(s, 9.4335, 3.75, 2.7433, 3.75);
	photo(s, 6.6902, 3.75, 2.7433, 3.75);
	photo(s, 1.1565, 0, 5.5102, 2.9303);

	navBar(s, { burgerX: 1.9565 });
	heading(s, 'Vision & ', 'Mission', { x: 7.3695, y: 1.3368, w: 4.1703 });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore', 7.3511, 2.3791, 4.9309, 0.6014);

	gradient(s, 1.1565, 2.9303, 5.5337, 4.5697, 18);
	[
		{ title: '1. Vision Title', ty: 3.7308, tw: 1.8656, x: 1.8831, by: 4.254 },
		{ title: '2. Mission Title', ty: 5.7252, tw: 2.1013, x: 1.8813, by: 6.2487 }
	].forEach(c => {
		text(s, c.title, { x: c.x, y: c.ty, w: c.tw, h: 0.4039, fontSize: 18, bold: true, color: WHITE, wrap: false });
		body(s, LOREM_VIS, c.x, c.by, 4.1088, 0.6014, { color: WHITE });
	});
}

/* --- 4. facilities ------------------------------------------------------- */
function slide04 (s) {
	box(s, 0, 4.5368, 6.6667, 2.9632, NAVY);
	navBar(s);

	heading(s, 'Best Our ', 'Facilities', { x: 1.0604, y: 1.3158, w: 4.7081 });
	body(s, LOREM_MED, 1.0665, 2.3791, 4.1124, 0.6014);

	[
		{ n: '01', bx: 7.3353, by: 1.4114, tx: 8.5373, ty: 1.323 },
		{ n: '02', bx: 7.3535, by: 3.362, tx: 8.5555, ty: 3.2735 },
		{ n: '03', bx: 7.3717, by: 5.3125, tx: 8.5737, ty: 5.2241 }
	].forEach(f => {
		text(s, 'Facilities', { x: f.tx, y: f.ty, w: 1.3397, h: 0.4039, fontSize: 18, bold: true, color: NAVY, wrap: false });
		body(s, LOREM_FAC, f.tx, f.ty + 0.3647, 3.6611, 0.6014);
		gradient(s, f.bx, f.by, 0.7893, 0.7893, 8);
		text(s, f.n, { x: f.bx, y: f.by, w: 0.7893, h: 0.7893, align: 'center', valign: 'middle', fontSize: 28, bold: true, color: WHITE });
	});

	photo(s, 1.1565, 3.75, 2.0959, 3.1391);
	photo(s, 3.5344, 3.75, 2.0959, 3.1391);
}

/* --- 5. projects --------------------------------------------------------- */
function slide05 (s) {
	navBar(s);
	heading(s, 'Best Our ', 'Project', { x: 8.1476, y: 1.3368, w: 4.1703 });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut', 8.1476, 2.3791, 4.1331, 0.6014);

	gradient(s, 3.5187, 3.7659, 8.6557, 2.3968, 20);
	[
		{ title: '1. Project Title', tx: 4.1771, ty: 4.3903, tw: 2.0058, bx: 4.1932, by: 4.8324 },
		{ title: '2. Project Title', tx: 8.1194, ty: 4.3524, tw: 2.4218, bx: 8.1355, by: 4.7944 }
	].forEach(c => {
		text(s, c.title, { x: c.tx, y: c.ty, w: c.tw, h: 0.4039, fontSize: 18, bold: true, color: WHITE });
		body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod', c.bx, c.by, 3.3903, 0.6014, { color: WHITE });
	});

	photo(s, 0, 1.4063, 2.7048, 2.1766);
	photo(s, 0, 3.9861, 2.7048, 2.1766);
	photo(s, 3.5187, 1.4063, 3.9483, 2.3593);
}

/* --- 6. team ------------------------------------------------------------- */
function slide06 (s) {
	navBar(s);
	heading(s, 'Meet Our ', 'Team', { x: 4.5815, y: 1.3368, w: 4.1703, align: 'center' });

	[
		{ name: 'Benjamin Shah', x: 1.1565 },
		{ name: 'Adora Montminy', x: 5.1711 },
		{ name: 'Alfredo Torres', x: 9.1857 }
	].forEach(m => {
		photo(s, m.x, 2.4949, 2.9908, 2.9089);
		gradient(s, m.x, 5.4038, 2.9908, 0.7072, 10);
		text(s, m.name, { x: m.x, y: 5.4038, w: 2.9908, h: 0.7072, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: WHITE });
	});
}

/* --- 7. services --------------------------------------------------------- */
function slide07 (s) {
	photo(s, 1.1565, 3.75, 11.0157, 3.75);
	navBar(s);

	heading(s, 'Best Our ', 'Service', { x: 1.0604, y: 1.3158, w: 4.7081 });
	body(s, LOREM_MED, 1.0665, 2.3791, 4.1124, 0.6014);
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut', 8.1402, 1.3948, 4.1417, 0.6014);

	[
		{ title: '1. Service', tw: 1.3775, cy: 1.4014, ty: 1.7598, by: 2.1958, icon: 'login' },
		{ title: '2. Service', tw: 1.4197, cy: 3.5446, ty: 3.9038, by: 4.3391, icon: 'crosshair' }
	].forEach(c => {
		gradient(s, 6.7671, c.cy, 5.4062, 1.7868, 16);
		const ix = 7.7903, iy = c.cy + 0.589; // icon centre
		if (c.icon === 'login') { // arrow entering an open bracket
			box(s, ix - 0.17, iy - 0.05, 0.20, 0.10, WHITE);
			s.addShape('triangle', { x: ix - 0.05, y: iy - 0.11, w: 0.22, h: 0.22, rotate: 90, fill: { color: WHITE }, line: NOLINE });
			s.addShape('roundRect', { x: ix - 0.02, y: iy - 0.17, w: 0.19, h: 0.34, rectRadius: 0.05, fill: { type: 'none' }, line: { color: WHITE, width: 1.75 } });
		} else { // crosshair over concentric circles
			s.addShape('ellipse', { x: ix - 0.16, y: iy - 0.16, w: 0.32, h: 0.32, fill: { type: 'none' }, line: { color: WHITE, width: 2 } });
			s.addShape('ellipse', { x: ix - 0.06, y: iy - 0.06, w: 0.12, h: 0.12, fill: { color: WHITE }, line: NOLINE });
			[[0, -0.19], [0, 0.19], [-0.19, 0], [0.19, 0]].forEach(([dx, dy]) => {
				box(s, ix + dx - 0.05, iy + dy - 0.05, 0.10, 0.10, WHITE);
			});
		}
		text(s, c.title, { x: 8.0922, y: c.ty, w: c.tw, h: 0.4039, fontSize: 18, bold: true, color: WHITE, wrap: false });
		body(s, 'Lorem ipsum dolor sit amet, consectetur adi piscing elit, sed consectetur adi piscing elit,', 7.4968, c.by, 3.9511, 0.6014, { color: WHITE });
	});
}

/* --- 8. portfolio -------------------------------------------------------- */
function slide08 (s) {
	photo(s, 7.4603, 2.6157, 4.7079, 3.4889);
	navBar(s);
	heading(s, 'Best Our ', 'Portfolio', { x: 8.1439, y: 1.3158, w: 4.7081 });
	box(s, 9.8137, 3.75, 2.3541, 2.3541, WHITE); // white gutter behind the small inset
	photo(s, 3.6979, 1.3958, 3.4176, 2.1975);
	photo(s, 10.1372, 4.0735, 2.0398, 2.0307);
	photo(s, 3.6979, 3.9069, 3.4176, 2.1975);
	photo(s, 1.1565, 3.9069, 2.1975, 2.1975);
	photo(s, 1.1565, 1.3958, 2.1975, 2.1975);
}

/* --- 9. pricing ---------------------------------------------------------- */
function slide09 (s) {
	// regular (green) card
	box(s, 9.4514, 1.8937, 3.0539, 3.7897, GREEN);
	s.addText('Register Now', { shape: 'rect', x: 10.0621, y: 4.7182, w: 1.8324, h: 0.6017, fill: { color: '808080', transparency: 90 }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 12, bold: true, color: WHITE, fontFace: FONT });
	s.addText('Regular Class', { shape: 'rect', x: 9.7908, y: 2.2664, w: 2.4361, h: 0.5709, fill: { color: '808080', transparency: 90 }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: WHITE, fontFace: FONT });
	text(s, [
		{ text: 'FREE ', options: { fontSize: 24, bold: true } },
		{ text: '/Month', options: { fontSize: 16, bold: true } }
	], { x: 9.8158, y: 2.9603, w: 2.3861, h: 0.5052, align: 'center', color: WHITE });
	['01', '02'].forEach((n, i) => {
		text(s, 'Your Service Title ' + n, { x: 10.3158, y: 3.7024 + i * 0.4299, w: 1.8861, h: 0.2777, fontSize: 10.5, color: WHITE });
		s.addShape('ellipse', { x: 10.1993, y: 3.7924 + i * 0.4299, w: 0.0972, h: 0.0972, fill: { color: WHITE }, line: NOLINE });
	});

	// premium (navy) card
	box(s, 5.7029, 1.4172, 3.7482, 4.7004, NAVY);
	s.addText('Register Now', { shape: 'rect', x: 6.3345, y: 5.0372, w: 2.5149, h: 0.7258, fill: { color: WHITE, transparency: 90 }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 16, bold: true, color: WHITE, fontFace: FONT });
	s.addText('Premium Class', { shape: 'rect', x: 6.1178, y: 1.8033, w: 2.9366, h: 0.6742, fill: { color: WHITE, transparency: 90 }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: WHITE, fontFace: FONT });
	text(s, [
		{ text: '$500 ', options: { fontSize: 28, bold: true } },
		{ text: '/Month', options: { fontSize: 16, bold: true } }
	], { x: 6.3095, y: 2.5804, w: 2.3861, h: 0.5717, align: 'center', color: WHITE });
	['01', '02', '03'].forEach((n, i) => {
		text(s, 'Your Service Title ' + n, { x: 6.8189, y: 3.4423 + i * 0.4645, w: 1.8861, h: 0.2777, fontSize: 10.5, color: WHITE });
		s.addShape('ellipse', { x: 6.7023, y: 3.5471 + i * 0.4645, w: 0.0972, h: 0.0972, fill: { color: WHITE }, line: NOLINE });
	});

	body(s, LOREM_MAEC, 1.0562, 2.9937, 3.4224, 1.2123, { fontSize: 11 });
	body(s, LOREM_MAEC, 1.0562, 4.6628, 3.4224, 1.2123, { fontSize: 11 });

	navBar(s);
	text(s, [
		{ text: 'Infographic ', options: { color: NAVY } },
		{ text: 'Section', options: { color: CYAN } }
	], { x: 1.0604, y: 1.3158, w: 2.9472, h: 1.3459, fontFace: DISPLAY, fontSize: 37 });
}

/* --- 10..18 share the centred "Infographic Section" title ---------------- */
function infographicTitle (s) {
	heading(s, 'Infographic ', 'Section', { x: 3.1111, y: 1.3158, w: 7.1111, align: 'center' });
}

/* --- 10. four numbered steps --------------------------------------------- */
function slide10 (s) {
	[
		{ n: '01', x: 0.9231, fill: NAVY },
		{ n: '02', x: 3.8254, fill: GREEN },
		{ n: '03', x: 6.7277, fill: NAVY },
		{ n: '04', x: 9.6300, fill: GREEN }
	].forEach(c => {
		text(s, c.n, { shape: 'rect', x: c.x + 0.9518, y: 3.9358, w: 0.9051, h: 0.9051, fill: { color: c.fill }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 18, bold: true, color: WHITE });
		text(s, 'Title Here', { x: c.x, y: 5.0296, w: 2.8086, h: 0.3029, align: 'center', fontSize: 12, bold: true });
		text(s, 'Lorem ipsum dolor sit amet, elit porttitor congue', { x: c.x, y: 5.3506, w: 2.8086, h: 0.6029, align: 'center', fontSize: 10.5, lineSpacingMultiple: 1.5 });
	});

	text(s, 'Lorem ipsum dolor sit amet, consect dolor sit amet, consect adipiscing elit. dolor dolor sit amet, consect dolor sit amet, consect adipiscing elit. dolor dolor sit amet, consect dolor sit amet, consect adipiscing elit. ',
		{ x: 2.2083, y: 2.747, w: 8.9167, h: 0.6029, align: 'center', fontSize: 10.5, lineSpacingMultiple: 1.5 });

	navBar(s);
	infographicTitle(s);
}

/* --- 11. twelve rule-marked notes ---------------------------------------- */
function slide11 (s) {
	const NOTE = 'Lorem ipsum dolor sit amet  consec tetur adla ipsum do Sit amet';
	[1.4, 4.3724, 7.3449, 10.3173].forEach(x => {
		[2.5208, 3.8796, 5.2385].forEach(y => {
			box(s, x - 0.2454, y + 0.1306, 0.0714, 0.4297, NAVY);
			body(s, NOTE, x, y, 1.9167, 1.1823, { fontSize: 11, color: INK_DK });
		});
	});
	navBar(s);
	infographicTitle(s);
}

/* --- 12. folded book panels ---------------------------------------------- */
function slide12 (s) {
	const PANEL = { w: 1.3476, h: 2.4459, points: parallelogramPoints(1.3476, 2.4459, 20526) };
	const rows = [
		{ back: CYAN_D50, front: CYAN_L40, y: 2.1017, fy: 2.0739, label: INK },
		{ back: NAVY_75, front: NAVY, y: 3.1737, fy: 3.1459, label: WHITE },
		{ back: GREEN_D75, front: GREEN, y: 4.2457, fy: 4.2179, label: WHITE }
	];
	const panel = (x, y, color, right, shadow) => s.addShape('custGeom', {
		x, y, w: PANEL.w, h: PANEL.h, points: PANEL.points,
		rotate: right ? 90 : 270, flipH: !!right, fill: { color }, line: NOLINE, shadow
	});
	// spine-shaded backs, then the two facing pages
	rows.forEach(r => panel(1.9425, r.y, r.back, false));
	rows.forEach(r => panel(4.5254, r.y, r.back, true));
	rows.forEach((r, i) => {
		panel(1.6979, r.fy, r.front, false, i === 2 ? shadow(4, 3, 0.4) : undefined);
		panel(4.7700, r.fy, r.front, true, i === 2 ? shadow(4, 3, 0.4) : undefined);
	});
	rows.forEach((r, i) => {
		const y = 3.1459 + i * 1.072;
		text(s, 'YOUR TITLE HERE', { x: 4.4029, y, w: 2.0817, h: 0.3029, rotate: -6.35, align: 'center', fontSize: 12, bold: true, color: r.label });
		text(s, 'YOUR TITLE HERE', { x: 1.3313, y, w: 2.0817, h: 0.3029, rotate: 6.35, flipH: true, align: 'center', fontSize: 12, bold: true, color: r.label });
	});

	// right hand copy
	[
		{ y: 2.5316, rule: CYAN_L40 },
		{ y: 3.8816, rule: NAVY },
		{ y: 5.2316, rule: GREEN }
	].forEach(c => {
		box(s, 7.5992, c.y + 0.1306, 0.0714, 0.4297, c.rule);
		body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, magna sedipsum dolor sit a dolor sit amet, consectetuer', 7.8446, c.y, 4.4753, 0.9042, { fontSize: 11 });
	});

	navBar(s);
	infographicTitle(s);
}

/* --- 13. arrow process --------------------------------------------------- */
function slide13 (s) {
	const poly = (x, y, w, h, pts, fill) => s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color: fill }, line: NOLINE });

	poly(1.1848, 4.3806, 3.7506, 1.3706,
		[{ x: 0, y: 0 }, { x: 2.769, y: 0 }, { x: 3.751, y: 0 }, { x: 3.264, y: 0.772 }, { x: 2.769, y: 0.772 }, { x: 2.769, y: 1.371 }, { x: 0, y: 1.371 }, { close: true }], GREEN_D75);
	poly(1.1848, 2.9828, 3.7593, 1.3983,
		[{ x: 0, y: 0 }, { x: 2.769, y: 0 }, { x: 2.769, y: 0.598 }, { x: 3.264, y: 0.598 }, { x: 3.759, y: 1.385 }, { x: 3.751, y: 1.398 }, { x: 0, y: 1.398 }, { close: true }], GREEN);

	box(s, 2.0294, 3.8272, 1.08, 1.08, WHITE, { shadow: shadow(17, 5) });
	s.addShape('rect', { x: 1.6252, y: 3.4229, w: 1.888, h: 1.888, fill: { type: 'none' }, line: { color: WHITE, width: 1 } });
	s.addShape('ellipse', { x: 2.5222, y: 4.2163, w: 0.3029, h: 0.3029, fill: { color: GREEN }, line: NOLINE });
	s.addShape('trapezoid', { x: 2.2952, y: 4.2439, w: 0.2806, h: 0.2456, rotate: 90, fill: { color: GREEN }, line: NOLINE });

	// chevrons: bottom row then top row
	const CHEV_DN = [{ x: 0.501, y: 0 }, { x: 1.563, y: 0 }, { x: 1.063, y: 0.772 }, { x: 0, y: 0.772 }, { close: true }];
	const CHEV_UP = [{ x: 0, y: 0 }, { x: 1.063, y: 0 }, { x: 1.572, y: 0.786 }, { x: 1.563, y: 0.8 }, { x: 0.501, y: 0.8 }, { x: 0.51, y: 0.786 }, { close: true }];
	[[4.5406, NAVY_D75], [5.6976, GREEN_D75]].forEach(([x, c]) => poly(x, 4.3806, 1.5631, 0.7723, CHEV_DN, c));
	poly(6.8519, 4.3806, 1.6191, 0.7723, [{ x: 0.501, y: 0 }, { x: 1.619, y: 0 }, { x: 1.385, y: 0.772 }, { x: 0, y: 0.772 }, { close: true }], NAVY_D75);
	[[4.5406, NAVY], [5.6976, GREEN]].forEach(([x, c]) => poly(x, 3.5806, 1.5719, 0.8, CHEV_UP, c));
	poly(6.8519, 3.5806, 1.6244, 0.8, [{ x: 0, y: 0 }, { x: 1.385, y: 0 }, { x: 1.624, y: 0.786 }, { x: 1.619, y: 0.8 }, { x: 0.501, y: 0.8 }, { x: 0.51, y: 0.786 }, { close: true }], NAVY);

	// little castle silhouette on the last chevron
	box(s, 7.4653, 3.3706, 0.2113, 0.2098, NAVY);
	box(s, 7.0799, 3.2447, 0.2113, 0.3357, NAVY);
	box(s, 7.1581, 3.1106, 0.1791, 0.3348, NAVY, { rotate: 90 });
	box(s, 7.8503, 3.2447, 0.2113, 0.3357, NAVY);
	box(s, 7.8038, 3.1106, 0.1791, 0.3348, NAVY, { rotate: 90 });

	text(s, 'Title Here', { x: 4.4653, y: 2.8087, w: 1.1913, h: 0.3029, align: 'center', fontSize: 12, bold: true, color: NAVY });
	text(s, 'Lorem ipsum dolor', { x: 4.2619, y: 3.0281, w: 1.5988, h: 0.3382, align: 'center', fontSize: 10.5, italic: true, lineSpacingMultiple: 1.5 });
	text(s, 'Title Here', { x: 5.7222, y: 5.3756, w: 1.1913, h: 0.3029, align: 'center', fontSize: 12, bold: true, color: GREEN });
	text(s, 'Lorem ipsum dolor', { x: 5.5181, y: 5.5949, w: 1.5988, h: 0.3382, align: 'center', fontSize: 10.5, italic: true, lineSpacingMultiple: 1.5 });

	[[2.7461, NAVY], [3.9369, NAVY_D75], [5.1272, NAVY_D50]].forEach(([y, c]) => {
		s.addShape('ellipse', { x: 9.2674, y: y + 0.09, w: 0.2387, h: 0.2387, fill: { color: c }, line: NOLINE });
		glyph(s, '\u2713', 9.2674, y + 0.09, 0.2387, WHITE, 9);
		body(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. ', 9.6069, y, 2.6687, 0.868);
	});

	navBar(s);
	infographicTitle(s);
}

/* --- 14. year cards ------------------------------------------------------ */
function slide14 (s) {
	const NOTE = 'Lorem ipsum dolor et dolor sit et amet, consectetuer adipiscing elit. dolor Maecenas porttmassa. Fusce dolor posuere, magna sedip sum dolor a dolor';
	[
		{ year: '2022', tab: 'a', x: 1.1837, dark: false },
		{ year: '2023', tab: 'b', x: 4.0210, dark: true },
		{ year: '2024', tab: 'c', x: 6.8583, dark: false },
		{ year: '2025', tab: 'd', x: 9.6956, dark: true }
	].forEach(c => {
		const fg = c.dark ? WHITE : NAVY;
		box(s, c.x, 2.6641, 2.4829, 3.4381, c.dark ? NAVY : WHITE, c.dark ? {} : { shadow: shadow(50, 20) });
		s.addShape('line', { x: c.x + 0.2409, y: 3.2081, w: 0, h: 2.8942, line: { color: fg, width: 1, dashType: 'dash' } });
		box(s, c.x + 0.1608, 3.0481, 0.16, 0.16, fg);
		text(s, c.year, { x: c.x + 0.3973, y: 3.2664, w: 1.4977, h: 0.37, fontSize: 16, bold: true, color: fg });
		body(s, NOTE, c.x + 0.3973, 3.6253, 1.8447, 2.1934, { color: c.dark ? WHITE : INK });
		text(s, c.tab, { shape: 'rect', x: c.x + 1.8249, y: 5.8005, w: 0.5143, h: 0.5143, fill: { color: c.dark ? WHITE : NAVY }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 14, bold: true, color: c.dark ? NAVY : WHITE, shadow: shadow(30, 15) });
	});
	navBar(s);
	infographicTitle(s);
}

/* --- 15. heptagon chain -------------------------------------------------- */
function slide15 (s) {
	const RIBBON = [
		{ x: 0.0049, y: 4.3934, w: 1.1822, h: 1.5721, pts: [{ x: 0, y: 0 }, { x: 1.182, y: 1.182 }, { x: 0.792, y: 1.572 }, { x: 0, y: 0.78 }, { close: true }] },
		{ x: 12.1424, y: 2.8446, w: 1.1909, h: 1.5808, pts: [{ x: 0.39, y: 0 }, { x: 1.191, y: 0.801 }, { x: 1.191, y: 1.581 }, { x: 0, y: 0.39 }, { close: true }] }
	];
	s.addShape('custGeom', Object.assign({ points: RIBBON[0].pts, fill: { color: GREEN, transparency: 20 }, line: NOLINE }, { x: RIBBON[0].x, y: RIBBON[0].y, w: RIBBON[0].w, h: RIBBON[0].h }));

	const NODES = [
		{ x: 0.9558, ribbon: NAVY, rot: 52.5, ring: GREEN, icon: 'doc', above: false },
		{ x: 3.3069, ribbon: GREEN, rot: 307.3, flip: true, ring: NAVY, icon: 'laptop', above: true },
		{ x: 5.6461, ribbon: NAVY, rot: 51.7, ring: GREEN, icon: 'person', above: false },
		{ x: 7.9970, ribbon: GREEN, rot: 308.6, flip: true, ring: NAVY, icon: 'target', above: true },
		{ x: 10.3355, ribbon: NAVY, rot: 52.5, ring: GREEN, icon: 'chat', above: false }
	];
	NODES.forEach(n => {
		s.addShape('rect', {
			x: n.x + 0.745, y: 2.4735, w: 0.552, h: 3.8625,
			rotate: n.flip ? 315 : 45, flipH: !!n.flip,
			fill: { color: n.ribbon, transparency: 20 }, line: NOLINE
		});
	});
	NODES.forEach(n => {
		s.addShape('heptagon', { x: n.x, y: 3.3849, w: 2.0417, h: 2.0417, rotate: n.rot, flipH: !!n.flip, fill: { color: n.ring }, line: NOLINE, shadow: shadow(25, 7) });
		s.addShape('heptagon', { x: n.x + 0.2083, y: 3.5932, w: 1.625, h: 1.625, rotate: n.rot, flipH: !!n.flip, fill: { color: WHITE }, line: NOLINE });
	});
	s.addShape('custGeom', Object.assign({ points: RIBBON[1].pts, fill: { color: GREEN, transparency: 20 }, line: NOLINE }, { x: RIBBON[1].x, y: RIBBON[1].y, w: RIBBON[1].w, h: RIBBON[1].h }));

	// icons inside the white heptagons
	NODES.forEach(n => {
		const cx = n.x + 1.0208, cy = 4.4058, c = n.ring === GREEN ? NAVY : GREEN;
		if (n.icon === 'doc') { // sheet with ruled lines and a magnifier
			box(s, cx - 0.24, cy - 0.30, 0.44, 0.58, c);
			[0, 1, 2].forEach(i => box(s, cx - 0.19, cy - 0.06 + i * 0.10, 0.30, 0.045, WHITE));
			s.addShape('ellipse', { x: cx - 0.03, y: cy - 0.02, w: 0.26, h: 0.26, fill: { color: WHITE }, line: { color: c, width: 2 } });
			box(s, cx + 0.17, cy + 0.19, 0.11, 0.07, c, { rotate: 45 });
		} else if (n.icon === 'laptop') { // screen on a stand
			box(s, cx - 0.31, cy - 0.28, 0.62, 0.42, c);
			box(s, cx - 0.24, cy - 0.21, 0.48, 0.28, WHITE);
			s.addShape('ellipse', { x: cx - 0.08, y: cy - 0.16, w: 0.16, h: 0.16, fill: { color: WHITE }, line: { color: c, width: 2 } });
			box(s, cx + 0.03, cy - 0.04, 0.08, 0.05, c, { rotate: 45 });
			box(s, cx - 0.05, cy + 0.14, 0.10, 0.10, c);
			box(s, cx - 0.40, cy + 0.22, 0.80, 0.08, c);
		} else if (n.icon === 'target') { // large bullseye with a dart
			target(s, cx, cy + 0.03, 0.29, c);
			dart(s, cx, cy + 0.03, 0.34, c);
		} else { // bust: head (bullseye or speech bubble) over shoulders
			if (n.icon === 'chat') {
				s.addShape('ellipse', { x: cx - 0.26, y: cy - 0.36, w: 0.52, h: 0.40, fill: { color: c }, line: NOLINE });
				s.addShape('triangle', { x: cx - 0.22, y: cy - 0.06, w: 0.14, h: 0.14, rotate: 200, fill: { color: c }, line: NOLINE });
				[-0.13, 0, 0.13].forEach(d => s.addShape('ellipse', { x: cx + d - 0.038, y: cy - 0.20, w: 0.076, h: 0.076, fill: { color: WHITE }, line: NOLINE }));
			} else {
				target(s, cx, cy - 0.17, 0.20, c);
				dart(s, cx, cy - 0.17, 0.26, c);
			}
			s.addShape('trapezoid', { x: cx - 0.35, y: cy + 0.05, w: 0.70, h: 0.32, fill: { color: c }, line: NOLINE });
			s.addShape('triangle', { x: cx - 0.11, y: cy + 0.05, w: 0.22, h: 0.17, flipV: true, fill: { color: WHITE }, line: NOLINE });
			s.addShape('triangle', { x: cx - 0.04, y: cy + 0.13, w: 0.08, h: 0.20, flipV: true, fill: { color: WHITE }, line: NOLINE });
		}
	});

	[
		{ x: 1.884, tx: 1.5271, y: 5.5942 },
		{ x: 4.3329, tx: 3.976, y: 2.5397 },
		{ x: 6.8929, tx: 6.5361, y: 5.5942 },
		{ x: 9.0100, tx: 8.6539, y: 2.5397 }
	].forEach(c => {
		text(s, 'Internal Documents', { x: c.x, y: c.y, w: 2.2528, h: 0.3029, align: 'center', fontSize: 12, bold: true });
		text(s, 'Lorem ipsum dolor sit amet, consec.', { x: c.tx, y: c.y + 0.2617, w: 2.9657, h: 0.3382, align: 'center', fontSize: 10.5, lineSpacingMultiple: 1.5 });
	});

	navBar(s);
	infographicTitle(s);
}

/* --- 16. arrow stack ----------------------------------------------------- */
function slide16 (s) {
	const OUTER = [{ x: 0, y: 0 }, { x: 0.974, y: 0 }, { x: 1.451, y: 0.563 }, { x: 0.974, y: 1.126 }, { x: 0, y: 1.126 },
		{ x: 0, y: 1.009 }, { x: 0.844, y: 1.009 }, { x: 1.222, y: 0.563 }, { x: 0.844, y: 0.117 }, { x: 0, y: 0.117 }, { close: true }];
	const INNER = [{ x: 0, y: 0 }, { x: 0.767, y: 0 }, { x: 1.087, y: 0.378 }, { x: 0.767, y: 0.755 }, { x: 0, y: 0.755 }, { close: true }];

	const ARROWS = [
		{ y: 2.6479, c: GREEN, flip: false, x: 5.4319, ix: 5.4319 },
		{ y: 3.8302, c: NAVY, flip: false, x: 5.4319, ix: 5.4319 },
		{ y: 5.0125, c: NAVY_60, flip: false, x: 5.4319, ix: 5.4319 },
		{ y: 3.2353, c: NAVY, flip: true, x: 6.4597, ix: 6.8236 },
		{ y: 4.4254, c: NAVY_60, flip: true, x: 6.4597, ix: 6.8236 }
	];
	ARROWS.forEach(a => {
		s.addShape('custGeom', { x: a.x, y: a.y, w: 1.4514, h: 1.1257, points: OUTER, flipH: a.flip, fill: { color: a.c }, line: NOLINE });
		s.addShape('custGeom', { x: a.ix, y: a.y + 0.1847, w: 1.0875, h: 0.7549, points: INNER, flipH: a.flip, fill: { color: WHITE }, line: NOLINE, shadow: shadow(7, 0, 0.25) });
		notebook(s, (a.flip ? 7.2583 : 5.7819) + 0.1515, a.y + 0.5625, 0.303, a.c);
	});

	[
		{ n: '01', c: GREEN, bx: 4.7701, by: 2.9663, tx: 3.3576, ty: 2.8773, align: 'right' },
		{ n: '02', c: NAVY, bx: 4.7701, by: 4.1231, tx: 3.3576, ty: 4.0492, align: 'right' },
		{ n: '03', c: NAVY_60, bx: 4.7701, by: 5.2808, tx: 3.3576, ty: 5.2049, align: 'right' },
		{ n: '04', c: NAVY, bx: 8.0243, by: 3.5279, tx: 8.7383, ty: 3.4492, align: 'left' },
		{ n: '05', c: NAVY_60, bx: 8.0243, by: 4.7431, tx: 8.7383, ty: 4.6601, align: 'left' }
	].forEach(c => {
		text(s, c.n, { shape: 'rect', x: c.bx, y: c.by, w: 0.5389, h: 0.5389, fill: { color: c.c }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 14, bold: true, color: WHITE });
		text(s, 'Title Here', { x: c.align === 'right' ? c.tx : c.tx, y: c.ty, w: 1.2097, h: 0.3029, align: c.align, fontSize: 12, bold: true, color: c.c });
		text(s, 'Lorem ipsum dolor sit amet, consectetuer adipis cing elit. Maecenas porttmassa. ',
			{ x: c.align === 'right' ? 0.6669 : 8.7383, y: c.ty + 0.2411, w: 3.9007, h: 0.6029, align: c.align === 'right' ? 'right' : 'justify', fontSize: 10.5, lineSpacingMultiple: 1.5 });
	});

	navBar(s);
	infographicTitle(s);
}

/* --- 17. triangle rosette ------------------------------------------------ */
function slide17 (s) {
	const PETALS = [
		{ rot: 0, lx: 5.9028, ly: 2.6479, tx: 5.9028, ty: 3.0312, band: NAVY, tri: NAVY_D75, flipH: false, icon: 'target' },
		{ rot: 60, lx: 7.2528, ly: 3.4544, tx: 6.5132, ty: 3.4107, band: GREEN, tri: GREEN_D75, flipH: false, icon: 'gift' },
		{ rot: 120, lx: 7.2757, ly: 5.0259, tx: 6.5368, ty: 4.1292, band: NAVY, tri: NAVY_D75, flipH: false, icon: 'card' },
		{ rot: 180, lx: 5.9028, ly: 5.7863, tx: 5.9028, ty: 4.4632, band: GREEN, tri: GREEN_D75, flipH: false, icon: 'target' },
		{ rot: 300, lx: 4.5451, ly: 3.4544, tx: 5.2840, ty: 3.4107, band: GREEN, tri: GREEN_D75, flipH: true, icon: 'card' },
		{ rot: 240, lx: 4.5222, ly: 5.0259, tx: 5.2611, ty: 4.1292, band: NAVY, tri: NAVY_D75, flipH: true, icon: 'gift' }
	];
	PETALS.forEach(p => {
		text(s, 'Title Here', {
			shape: 'rect', x: p.lx, y: p.ly, w: 1.5359, h: 0.3837, rotate: p.rot, flipH: p.flipH,
			fill: { color: p.band }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 11, bold: true, color: WHITE
		});
		s.addShape('triangle', { x: p.tx, y: p.ty, w: 1.5359, h: 1.3238, rotate: p.rot, flipV: true, flipH: p.flipH, fill: { color: p.tri }, line: NOLINE });
	});

	// icons at the middle of each triangle
	const R = 1.0, CX = 6.6707, CY = 4.4341;
	PETALS.forEach((p, i) => {
		const a = (-90 + i * 60) * Math.PI / 180;
		const cx = CX + R * 0.62 * Math.cos(a), cy = CY + R * 0.62 * Math.sin(a);
		if (p.icon === 'target') {
			target(s, cx, cy + 0.02, 0.19, WHITE, p.tri);
			dart(s, cx, cy + 0.02, 0.24, WHITE);
			return;
		}
		if (p.icon === 'gift') {
			box(s, cx - 0.19, cy - 0.10, 0.38, 0.29, WHITE);
			box(s, cx - 0.22, cy - 0.18, 0.44, 0.09, WHITE);
			box(s, cx - 0.035, cy - 0.18, 0.07, 0.37, p.tri);
			s.addShape('ellipse', { x: cx - 0.15, y: cy - 0.28, w: 0.13, h: 0.13, fill: { type: 'none' }, line: { color: WHITE, width: 2 } });
			s.addShape('ellipse', { x: cx + 0.02, y: cy - 0.28, w: 0.13, h: 0.13, fill: { type: 'none' }, line: { color: WHITE, width: 2 } });
			return;
		}
		box(s, cx - 0.20, cy - 0.15, 0.40, 0.30, WHITE); // wallet / card
		box(s, cx - 0.20, cy - 0.07, 0.40, 0.06, p.tri);
		box(s, cx + 0.06, cy + 0.01, 0.10, 0.07, p.tri);
	});

	[
		{ n: '01', c: NAVY, side: 'l', y: 2.5870, by: 2.7196 },
		{ n: '03', c: GREEN, side: 'l', y: 4.6976, by: 4.8306 },
		{ n: '02', c: NAVY, side: 'r', y: 2.5870, by: 2.7196 },
		{ n: '04', c: GREEN, side: 'r', y: 4.6976, by: 4.8306 }
	].forEach(c => {
		const left = c.side === 'l';
		text(s, 'Lorem ipsum dolor sit amet, ipsum dolor sit amet, ipsum dolor sit amet, ipsum dolor sit amet',
			{ x: left ? 1.8264 : 8.8208, y: c.y, w: 2.7090, h: 0.868, align: left ? 'justify' : 'right', fontSize: 10.5, lineSpacingMultiple: 1.5 });
		text(s, c.n, { shape: 'rect', x: left ? 1.1840 : 11.6459, y: c.by, w: 0.5257, h: 0.5257, fill: { color: c.c }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 12, bold: true, color: WHITE });
		const btnX = left ? 1.9222 : 10.4028;
		const tipX = left ? 2.9392 : 10.0799;
		text(s, 'Learn More', { shape: 'rect', x: btnX, y: c.by + 0.8877, w: 1.0296, h: 0.3357, fill: { color: c.c }, line: NOLINE, align: 'center', valign: 'middle', fontSize: 9, color: WHITE });
		box(s, tipX, c.by + 0.8877, 0.3357, 0.3357, WHITE, { shadow: shadow(30, 10) });
		s.addShape('custGeom', {
			x: tipX + 0.095, y: c.by + 1.0088, w: 0.146, h: 0.094,
			rotate: left ? 90 : 270, flipH: !left, fill: { color: c.c }, line: NOLINE,
			points: [{ x: 0.129, y: 0.092 }, { x: 0.122, y: 0.092 }, { x: 0.072, y: 0.044 }, { x: 0.024, y: 0.092 },
				{ x: 0.017, y: 0.092 }, { x: 0.002, y: 0.078 }, { x: 0.002, y: 0.070 }, { x: 0.069, y: 0.002 },
				{ x: 0.076, y: 0.002 }, { x: 0.143, y: 0.070 }, { x: 0.143, y: 0.078 }, { close: true }]
		});
	});

	navBar(s);
	infographicTitle(s);
}

/* --- 18. growth arrows + bars -------------------------------------------- */
function slide18 (s) {
	[
		{ x: 1.0827, y: 2.6479, w: 3.6472, h: 3.4542, a1: 52886, a2: 32237, c: NAVY_75 },
		{ x: 2.6893, y: 3.1438, w: 3.1226, h: 2.9583, a1: 50000, a2: 33651, c: NAVY_50 },
		{ x: 3.9840, y: 3.5170, w: 2.7292, h: 2.5851, a1: 50000, a2: 42903, c: GREEN }
	].forEach(a => {
		s.addShape('custGeom', { x: a.x, y: a.y, w: a.w, h: a.h, points: upArrowPoints(a.w, a.h, a.a1, a.a2), fill: { color: a.c }, line: NOLINE });
	});
	[['2024', 1.8757, 0.8299], ['2025', 3.4049, 0.8144], ['2026', 4.5880, 0.8232]].forEach(([y, x, w]) => {
		text(s, y, { x, y: 5.7357, w, h: 0.4039, fontSize: 18, bold: true, color: WHITE });
	});

	[
		{ label: 'Your Title Here 01', pct: '60%', y: 2.9515, fill: NAVY_75, bar: 2.3951, hx: 9.8563 },
		{ label: 'Your Title Here 02', pct: '72%', y: 3.8265, fill: NAVY_50, bar: 3.1006, hx: 10.5798 },
		{ label: 'Your Title Here 03', pct: '93%', y: 4.7015, fill: GREEN, bar: 3.5252, hx: 11.0073 }
	].forEach(r => {
		s.addShape('rect', { x: 7.5755, y: r.y, w: 3.9192, h: 0.1019, fill: { color: GREY_85, transparency: 50 }, line: NOLINE });
		box(s, 7.5755, r.y - 0.0164, r.bar, 0.1183, r.fill);
		box(s, r.hx, r.y - 0.0615, 0.1969, 0.1969, r.fill);
		text(s, r.label, { x: 7.4990, y: r.y - 0.4315, w: 2.5957, h: 0.345, fontSize: 12, lineSpacingMultiple: 1.3 });
		text(s, r.pct, { x: 11.7361, y: r.y - 0.1265, w: 0.7039, h: 0.321, fontSize: 11, lineSpacingMultiple: 1.3 });
	});

	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation',
		7.4617, 5.3233, 4.7891, 0.868);

	navBar(s);
	infographicTitle(s);
}

/* --- 19. contact --------------------------------------------------------- */
function slide19 (s) {
	photo(s, 1.1565, 3.75, 3.9358, 2.3553);
	photo(s, 9.2757, 3.75, 2.8934, 2.3553);
	photo(s, 9.2757, 1.3947, 2.8934, 2.3553);

	navBar(s);
	heading(s, 'Contact ', 'Us', { x: 1.0604, y: 1.3158, w: 2.9472 });
	body(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do sit amet', 1.0665, 2.3791, 3.3218, 0.6014);

	gradient(s, 5.0932, 1.3947, 4.1825, 3.9679, 18);
	[
		{ icon: 'phone', label: '+12345678910', w: 1.6333, y: 2.2288, cy: 2.3670 },
		{ icon: 'plane', label: 'yourmail@gmail.com', w: 1.9528, y: 2.9028, cy: 3.0410 },
		{ icon: 'globe', label: 'www.yourwebsite.com', w: 1.9528, y: 3.5768, cy: 3.7150 },
		{ icon: 'house', label: '123 Anywhere ST,, Any City', w: 2.2590, y: 4.2508, cy: 4.3890 }
	].forEach(r => {
		const cx = 5.8535;
		if (r.icon === 'phone') {
			s.addShape('roundRect', { x: cx - 0.13, y: r.cy - 0.13, w: 0.26, h: 0.26, rectRadius: 0.06, fill: { color: WHITE }, line: NOLINE });
			glyph(s, '\u260E', cx - 0.13, r.cy - 0.14, 0.26, mix(GREEN, CYAN, 0.4), 11);
		} else if (r.icon === 'plane') planeIcon(s, cx, r.cy, 0.13, WHITE);
		else if (r.icon === 'globe') globeIcon(s, cx, r.cy, 0.12, WHITE);
		else houseIcon(s, cx, r.cy, 0.13, WHITE, mix(GREEN, CYAN, 0.47));
		text(s, r.label, { x: 6.3718, y: r.y, w: r.w, h: 0.2777, fontSize: 10.5, color: WHITE });
	});
}

/* --- 20. thank you ------------------------------------------------------- */
function slide20 (s) {
	photo(s, 0, 1.4111, 8.2532, 4.7278);
	navBar(s);
	gradient(s, 8.5972, 1.4111, 3.5799, 2.3389, 14);

	text(s, [
		{ text: 'THANK', options: { color: WHITE } },
		{ text: ' YOU!', options: { color: NAVY } }
	], { x: 3.5250, y: 4.2833, w: 8.3708, h: 2.4229, fontFace: DISPLAY, fontSize: 138 });

	gradient(s, 4.3278, 1.4111, 3.1352, 2.3049, 12);
	text(s, 'See You Next Time', { x: 4.9722, y: 1.6882, w: 2.2433, h: 0.7072, fontSize: 18, bold: true, color: WHITE });
	body(s, LOREM_SHORT, 4.9840, 2.4951, 1.9158, 0.8665, { color: WHITE });

	photo(s, 8.7648, 1.5771, 3.2454, 2.0063);
}

/* ==================================================================== main */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
	slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15,
	slide16, slide17, slide18, slide19, slide20];

const pptx = new pptxgen();
pptx.defineLayout({ name: 'GOLF_16x9', width: 13.333, height: 7.5 });
pptx.layout = 'GOLF_16x9';
pptx.title = 'Golf Course - Sport Presentation Template';
pptx.theme = { headFontFace: DISPLAY, bodyFontFace: FONT };

BUILDERS.forEach(build => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '14c9bbcd-8994-4364-a295-4bd241fa7555_grok_final.pptx') })
	.then(f => console.log('wrote ' + f));
