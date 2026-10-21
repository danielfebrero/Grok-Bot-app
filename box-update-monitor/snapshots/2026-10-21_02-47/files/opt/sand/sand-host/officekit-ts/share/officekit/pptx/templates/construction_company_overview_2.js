/**
 * "kayuadam" — Construction Presentation (18 slides, 13.333in x 7.5in)
 * Standalone re-creation of the reference deck with pptxgenjs.
 *
 * Raster/photo content in the original lives in empty picture placeholders;
 * those are drawn here as outlined "[image]" placeholder frames.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const ORANGE = 'FFA73D'; // brand accent
const NAVY = '1D1A31'; // brand dark
const WHITE = 'FFFFFF';
const MUTED = 'C9C8D0'; // body copy on navy
const GRAY50 = '808080'; // bg1 lumMod 50%  – body copy on white
const GRAY65 = 'A6A6A6'; // bg1 lumMod 65%
const GRAY_59 = '595959'; // tx1 lumMod 65% / lumOff 35%
const GRAY_40 = '404040'; // tx1 lumMod 75% / lumOff 25%
const BLACK = '000000';
const TRACK = '4E4B66'; // unfilled part of a progress bar
const PH_LINE = 'E4E4E9'; // image-placeholder frame
const PH_TEXT = 'C9C8D0'; // image-placeholder caption

/* -------------------------------------------------------------------- fonts */
const SEMI = 'Poppins SemiBold';
const POPPINS = 'Poppins';
const LATO = 'Lato';
const WORK = 'Work Sans Medium';
const OPENS = 'Open Sans';

/* ------------------------------------------------------------ shared copy */
const SUB = 'Your Subtitle Here';
const EYEBROW_ABOUT = 'About Company';
const EYEBROW_KAYU = 'Kayuadam Company';

/* ------------------------------------------------------------------ helpers */

// Solid / outlined rectangle.
function rect(s, x, y, w, h, opts) {
	s.addShape('rect', Object.assign({ x, y, w, h }, opts));
}

// Text box. Reference text boxes are top-anchored and un-filled.
function text(s, str, opts) {
	s.addText(str, Object.assign({ valign: 'top', fontFace: LATO, fontSize: 18, color: BLACK }, opts));
}

// Small orange kicker above a section heading.
function eyebrow(s, x, y, w, str, fontSize, align) {
	text(s, str, {
		x, y, w, h: fontSize && fontSize <= 16 ? 0.37 : 0.404,
		fontFace: WORK, fontSize: fontSize || 18, color: ORANGE, align: align || 'left',
	});
}

// Large section heading.
function heading(s, x, y, w, str, fontSize, color, align) {
	const h = { 54: 1.01, 36: 0.707, 32: 0.64, 28: 0.572 }[fontSize] || 0.64;
	text(s, str, { x, y, w, h, fontFace: SEMI, fontSize, bold: true, color: color || NAVY, align: align || 'left' });
}

// Bold 12pt sub-heading.
function subhead(s, x, y, w, str, color, align) {
	text(s, str || SUB, { x, y, w, h: 0.303, fontFace: SEMI, fontSize: 12, bold: true, color: color || BLACK, align: align || 'left' });
}

// Justified 12pt body copy at 150% line spacing.
function body(s, x, y, w, h, str, color, align) {
	text(s, str, {
		x, y, w, h, fontFace: LATO, fontSize: 12, color: color || GRAY50,
		align: align || 'justify', lineSpacingMultiple: 1.5,
	});
}

// Uppercase tracked-out label ("LEARN MORE", "BUSINESS PROJECT", ...).
function label(s, x, y, w, h, str, color, align, fontFace, fontSize) {
	text(s, str, {
		x, y, w, h, fontFace: fontFace || POPPINS, fontSize: fontSize || 12,
		bold: true, charSpacing: 1, color: color || NAVY, align: align || 'left',
	});
}

// Thin arrow that follows a "LEARN MORE" label.
function arrow(s, x, y, color) {
	s.addShape('line', { x, y, w: 0.148, h: 0, line: { color: color || NAVY, width: 1.25, endArrowType: 'triangle' } });
}

/* ------------------------------------------------------------- icon toolkit
 * The reference deck uses flat line-art construction glyphs. Each one is
 * redrawn here as custom geometry expressed in normalised 0..1 coordinates
 * inside the glyph's own bounding box, so an icon can be dropped at any size.
 */
const STROKE = 3;

function pts(b, list, close) {
	const out = list.map(([px, py]) => ({ x: b.w * px, y: b.h * py }));
	if (close) out.push({ close: true });
	return out;
}
// Open / closed outline drawn with the icon colour.
function stroke(s, b, list, close) {
	s.addShape('custGeom', {
		x: b.x, y: b.y, w: b.w, h: b.h, points: pts(b, list, close),
		fill: { type: 'none' }, line: { color: b.c, width: STROKE },
	});
}
// Filled polygon.
function solid(s, b, list) {
	s.addShape('custGeom', {
		x: b.x, y: b.y, w: b.w, h: b.h, points: pts(b, list, true),
		fill: { color: b.c }, line: { type: 'none' },
	});
}
function ring(s, b, cx, cy, r, filled) {
	s.addShape('ellipse', {
		x: b.x + b.w * (cx - r), y: b.y + b.h * (cy - r), w: b.w * r * 2, h: b.h * r * 2,
		fill: filled ? { color: b.c } : { type: 'none' }, line: { color: b.c, width: STROKE },
	});
}
function arcOf(s, b, cx, cy, r, from, to) {
	s.addShape('arc', {
		x: b.x + b.w * (cx - r), y: b.y + b.h * (cy - r), w: b.w * r * 2, h: b.h * r * 2,
		angleRange: [from, to], fill: { type: 'none' }, line: { color: b.c, width: STROKE },
	});
}
function bounds(x, y, w, h, c) { return { x, y, w, h, c }; }

// traffic cone — slide 1
function iconCone(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.38, 0.04], [0.16, 0.84], [0.84, 0.84], [0.62, 0.04]], true);
	stroke(s, b, [[0.31, 0.30], [0.69, 0.30]]);
	stroke(s, b, [[0.24, 0.57], [0.76, 0.57]]);
	solid(s, b, [[0.02, 0.88], [0.98, 0.88], [0.98, 1], [0.02, 1]]);
}
// site plan with pencil — slide 3
function iconBlueprint(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.10, 0.02], [0.90, 0.02], [0.90, 0.60], [0.10, 0.60]], true);
	stroke(s, b, [[0.22, 0.28], [0.22, 0.50], [0.50, 0.50], [0.50, 0.28]]);
	stroke(s, b, [[0.16, 0.30], [0.36, 0.14], [0.56, 0.30]]);
	stroke(s, b, [[0.60, 0.32], [0.80, 0.32], [0.80, 0.50], [0.60, 0.50]], true);
	stroke(s, b, [[0.04, 0.74], [0.70, 0.74]]);
	stroke(s, b, [[0.04, 0.92], [0.55, 0.92]]);
	solid(s, b, [[0.72, 0.66], [0.98, 0.80], [0.72, 0.94]]);
}
// power drill firing at a wall — slide 5
function iconDrill(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.02, 0.08], [0.46, 0.08], [0.46, 0.44], [0.26, 0.44], [0.22, 1], [0.06, 1], [0.02, 0.44]], true);
	stroke(s, b, [[0.10, 0.18], [0.26, 0.18], [0.26, 0.30], [0.10, 0.30]], true);
	stroke(s, b, [[0.46, 0.14], [0.62, 0.14], [0.62, 0.36], [0.46, 0.36]], true);
	stroke(s, b, [[0.62, 0.22], [0.78, 0.22], [0.78, 0.28], [0.62, 0.28]], true);
	stroke(s, b, [[0.90, 0], [0.90, 0.62]]);
	[[0.80, 0.10], [0.82, 0.25], [0.82, 0.40]].forEach(([fx, fy]) => stroke(s, b, [[fx, fy], [fx + 0.14, fy - 0.04]]));
}
// clipboard checklist — slide 5
function iconClipboard(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.04, 0.08], [0.96, 0.08], [0.96, 1], [0.04, 1]], true);
	solid(s, b, [[0.32, 0], [0.68, 0], [0.68, 0.14], [0.32, 0.14]]);
	stroke(s, b, [[0.36, 0.22], [0.64, 0.22], [0.64, 0.38], [0.50, 0.48], [0.36, 0.38]], true);
	[0.60, 0.72, 0.84].forEach((fy) => {
		stroke(s, b, [[0.16, fy - 0.03], [0.22, fy + 0.02], [0.30, fy - 0.06]]);
		stroke(s, b, [[0.38, fy], [0.84, fy]]);
	});
}
// bulldozer — slide 8
function iconBulldozer(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.26, 0.34], [0.30, 0.04], [0.60, 0.04], [0.62, 0.34]], true);
	stroke(s, b, [[0.62, 0.16], [0.86, 0.16], [0.86, 0.34]]);
	solid(s, b, [[0.06, 0.38], [0.86, 0.38], [0.86, 0.48], [0.06, 0.48]]);
	stroke(s, b, [[0.10, 0.56], [0.74, 0.56], [0.80, 0.98], [0.06, 0.98]], true);
	ring(s, b, 0.24, 0.78, 0.15);
	ring(s, b, 0.58, 0.78, 0.15);
	stroke(s, b, [[0.94, 0.28], [0.94, 0.98]]);
}
// tower crane — slide 8
function iconCrane(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.30, 0.14], [0.30, 0.86]]);
	stroke(s, b, [[0.42, 0.14], [0.42, 0.86]]);
	stroke(s, b, [[0.30, 0.36], [0.42, 0.52], [0.30, 0.68], [0.42, 0.84]]);
	stroke(s, b, [[0.02, 0.14], [0.98, 0.14]]);
	stroke(s, b, [[0.36, 0], [0.90, 0.14]]);
	stroke(s, b, [[0.06, 0.14], [0.36, 0]]);
	stroke(s, b, [[0.76, 0.14], [0.76, 0.44]]);
	stroke(s, b, [[0.62, 0.44], [0.90, 0.44], [0.86, 0.58], [0.66, 0.58]], true);
	stroke(s, b, [[0.12, 0.86], [0.60, 0.86], [0.66, 1], [0.06, 1]], true);
}
// work boot — slide 8
function iconBoot(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.20, 0.02], [0.52, 0.02], [0.56, 0.52], [0.92, 0.62], [0.96, 0.86], [0.06, 0.86], [0.16, 0.40]], true);
	stroke(s, b, [[0.24, 0.26], [0.54, 0.26]]);
	stroke(s, b, [[0.58, 0.62], [0.92, 0.70]]);
	solid(s, b, [[0.04, 0.88], [0.98, 0.88], [0.98, 1], [0.04, 1]]);
}
// power plug and cable — slide 8
function iconPlug(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.06, 0.04], [0.24, 0.22]]);
	stroke(s, b, [[0.30, 0], [0.48, 0.18]]);
	stroke(s, b, [[0.10, 0.26], [0.36, 0], [0.60, 0.24], [0.34, 0.50]], true);
	stroke(s, b, [[0.52, 0.34], [0.76, 0.58]]);
	arcOf(s, b, 0.66, 0.72, 0.26, 200, 100);
	ring(s, b, 0.74, 0.80, 0.10);
}
// wheelbarrow — slide 9
function iconWheelbarrow(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	solid(s, b, [[0.22, 0.18], [0.94, 0.18], [0.80, 0.54], [0.34, 0.54]]);
	stroke(s, b, [[0.22, 0.18], [0.94, 0.18], [0.80, 0.54], [0.34, 0.54]], true);
	stroke(s, b, [[0.16, 0.10], [0.16, 0.34]]);
	stroke(s, b, [[0.34, 0.54], [0.86, 0.80]]);
	stroke(s, b, [[0.80, 0.54], [0.56, 0.86]]);
	ring(s, b, 0.28, 0.74, 0.20);
	ring(s, b, 0.82, 0.84, 0.11);
}
// hammer striking a post — slide 9
function iconHammer(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.04, 0], [0.04, 1]]);
	solid(s, b, [[0.22, 0.22], [0.44, 0.04], [0.62, 0.26], [0.56, 0.34], [0.66, 0.46], [0.52, 0.56], [0.42, 0.44], [0.34, 0.50]]);
	solid(s, b, [[0.52, 0.52], [0.66, 0.62], [0.92, 1], [0.74, 1]]);
}
// screwdriver — slide 9
function iconScrewdriver(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	solid(s, b, [[0.60, 0], [0.86, 0.16], [0.62, 0.42], [0.44, 0.28]]);
	stroke(s, b, [[0.44, 0.28], [0.62, 0.42], [0.24, 0.84], [0.06, 0.70]], true);
	stroke(s, b, [[0.06, 0.70], [0.24, 0.84], [0.10, 1], [0, 0.88]], true);
}
// high-visibility vest — slide 10
function iconVest(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.30, 0.02], [0.04, 0.16], [0.04, 0.98], [0.96, 0.98], [0.96, 0.16], [0.70, 0.02], [0.50, 0.30]], true);
	stroke(s, b, [[0.30, 0.02], [0.50, 0.30]]);
	stroke(s, b, [[0.50, 0.30], [0.50, 0.98]]);
	[0.44, 0.60].forEach((fy) => stroke(s, b, [[0.04, fy], [0.96, fy]]));
	[0.22, 0.78].forEach((fx) => stroke(s, b, [[fx, 0.12], [fx, 0.98]]));
}
// jackhammer — slide 10
function iconJackhammer(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.30, 0], [0.70, 0], [0.70, 0.26], [0.30, 0.26]], true);
	stroke(s, b, [[0.16, 0.08], [0.30, 0.08]]);
	stroke(s, b, [[0.70, 0.08], [0.86, 0.08]]);
	stroke(s, b, [[0.42, 0.26], [0.58, 0.26], [0.58, 0.52], [0.42, 0.52]], true);
	solid(s, b, [[0.44, 0.52], [0.56, 0.52], [0.52, 0.74], [0.48, 0.74]]);
	stroke(s, b, [[0.04, 0.86], [0.96, 0.86]]);
	[[0.10, 1], [0.30, 1], [0.70, 1], [0.90, 1]].forEach(([fx, fy]) => stroke(s, b, [[0.50, 0.80], [fx, fy]]));
}
// hard hat with goggles — slide 10
function iconHelmet(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.16, 0.44], [0.22, 0.14], [0.36, 0.02], [0.64, 0.02], [0.78, 0.14], [0.84, 0.44]], true);
	[0.34, 0.50, 0.66].forEach((fx) => stroke(s, b, [[fx, 0.06], [fx, 0.44]]));
	solid(s, b, [[0.02, 0.46], [0.98, 0.46], [0.98, 0.56], [0.02, 0.56]]);
	stroke(s, b, [[0.14, 0.64], [0.86, 0.64], [0.86, 0.92], [0.14, 0.92]], true);
	stroke(s, b, [[0.50, 0.64], [0.50, 0.92]]);
}
// map pin — slide 17
function iconPin(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	s.addShape('teardrop', {
		x: x + w * 0.04, y, w: w * 0.92, h: h * 0.86, rotate: 135,
		fill: { type: 'none' }, line: { color: c, width: STROKE },
	});
	ring(s, b, 0.50, 0.38, 0.20);
}
// clock — slide 17
function iconClock(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	ring(s, b, 0.5, 0.5, 0.5);
	stroke(s, b, [[0.50, 0.22], [0.50, 0.52], [0.72, 0.62]]);
}
// telephone handset with signal waves — slide 17
function iconPhone(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	s.addShape('moon', {
		x: x + w * 0.02, y: y + h * 0.2, w: w * 0.62, h: h * 0.8, rotate: 315,
		fill: { color: c }, line: { type: 'none' },
	});
	arcOf(s, b, 0.60, 0.40, 0.30, 285, 55);
	arcOf(s, b, 0.60, 0.40, 0.48, 285, 55);
}
// envelope — slide 17
function iconMail(s, x, y, w, h, c) {
	const b = bounds(x, y, w, h, c);
	stroke(s, b, [[0.02, 0.10], [0.98, 0.10], [0.98, 0.90], [0.02, 0.90]], true);
	stroke(s, b, [[0.02, 0.10], [0.50, 0.55], [0.98, 0.10]]);
	stroke(s, b, [[0.02, 0.90], [0.38, 0.52]]);
	stroke(s, b, [[0.98, 0.90], [0.62, 0.52]]);
}

// Tiny social glyphs (linkedin / instagram / facebook / twitter) — slide 6
function socialRow(s, x, y) {
	const li = bounds(x, y + 0.005, 0.122, 0.112, ORANGE);
	solid(s, li, [[0, 0.30], [0.24, 0.30], [0.24, 1], [0, 1]]);
	solid(s, li, [[0, 0], [0.24, 0], [0.24, 0.20], [0, 0.20]]);
	solid(s, li, [[0.36, 0.30], [1, 0.30], [1, 1], [0.76, 1], [0.76, 0.55], [0.60, 0.55], [0.60, 1], [0.36, 1]]);
	s.addShape('roundRect', {
		x: x + 0.301, y, w: 0.135, h: 0.135, rectRadius: 0.045,
		fill: { type: 'none' }, line: { color: ORANGE, width: 2 },
	});
	s.addShape('ellipse', {
		x: x + 0.34, y: y + 0.039, w: 0.057, h: 0.057,
		fill: { type: 'none' }, line: { color: ORANGE, width: 2 },
	});
	const fb = bounds(x + 0.614, y - 0.002, 0.079, 0.145, ORANGE);
	solid(s, fb, [[0.35, 0.28], [1, 0.28], [1, 0], [0.62, 0], [0.35, 0.14]]);
	solid(s, fb, [[0.35, 0.28], [0.72, 0.28], [0.72, 1], [0.35, 1]]);
	solid(s, fb, [[0, 0.36], [0.90, 0.36], [0.90, 0.56], [0, 0.56]]);
	const tw = bounds(x + 0.871, y + 0.009, 0.135, 0.11, ORANGE);
	solid(s, tw, [[0, 0.72], [0.35, 0.95], [0.80, 0.65], [1, 0.16], [0.82, 0.24], [0.55, 0], [0.28, 0.30], [0.10, 0.20], [0.20, 0.50]]);
}

// Empty picture placeholder from the original layout.
function imageBox(s, x, y, w, h) {
	s.addShape('rect', { x, y, w, h, fill: { type: 'none' }, line: { color: PH_LINE, width: 0.75, dashType: 'dash' } });
	s.addText('[image]', {
		x, y: y + h / 2 - 0.16, w, h: 0.32,
		fontFace: LATO, fontSize: 10, color: PH_TEXT, align: 'center', valign: 'middle',
	});
}

// Horizontal accent rule.
function rule(s, x, y, w, color, width) {
	s.addShape('line', { x, y, w, h: 0, line: { color, width: width || 1.5 } });
}

/* ------------------------------------------------------------------- slides */

// 1 — title
function slide01(pres) {
	const s = pres.addSlide();
	rect(s, 5.014, 5.333, 8.319, 2.167, { fill: { color: ORANGE } });
	text(s, 'kayuadam', { x: 7.126, y: 2.535, w: 5.58, h: 1.01, fontFace: SEMI, fontSize: 54, bold: true, color: NAVY });
	rect(s, 0, 0, 1.261, 1.101, { fill: { color: NAVY } });
	text(s, '2023', { x: 0.231, y: 6.778, w: 1.388, h: 0.37, fontFace: SEMI, fontSize: 16, bold: true, color: GRAY50 });
	rule(s, 1.261, 6.949, 3.346, GRAY50, 1.75);
	text(s, 'Construction Presentation', { x: 7.126, y: 3.545, w: 5.58, h: 0.404, fontFace: POPPINS, color: GRAY50 });
	rect(s, 7.244, 4.216, 1.581, 0.439, { fill: { type: 'none' }, line: { color: ORANGE, width: 1.25 } });
	label(s, 7.244, 4.291, 1.581, 0.303, 'LEARN MORE', ORANGE, 'center');
	iconCone(s, 0.394, 0.315, 0.473, 0.472, ORANGE);
	imageBox(s, 0, 1.101, 6.475, 5.297);
}

// 2 — introduction, 2x2 copy grid on a navy panel
function slide02(pres) {
	const s = pres.addSlide();
	rect(s, 2.691, 2.361, 10.642, 4.793, { fill: { color: NAVY } });
	imageBox(s, 0, 3.757, 4.652, 3.397);
	rect(s, 4.331, 4.941, 0.642, 2.213, { fill: { color: ORANGE } });
	s.addText('CONSTRUCTION', {
		x: 4.45, y: 5.229, w: 0.404, h: 1.623, vert: 'vert270', wrap: false, valign: 'top',
		fontFace: POPPINS, fontSize: 12, bold: true, charSpacing: 1, color: NAVY,
	});
	heading(s, 7.279, 1.378, 5.45, 'Introduction Kayuadam', 28);
	eyebrow(s, 7.279, 0.974, 5.45, EYEBROW_ABOUT);
	body(s, 7.279, 2.788, 5.45, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat sed lectus vestibulum mattis ullamcorper. Nulla facilisi', MUTED);

	const cell = 'Bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat';
	[[5.574, 4.244], [9.391, 4.244], [5.574, 5.699], [9.391, 5.699]].forEach(([x, y]) => {
		subhead(s, x, y, 3.338, SUB, WHITE, 'justify');
		body(s, x, y + 0.302, 3.338, 0.667, cell, MUTED);
	});
}

// 3 — about our business, hexagon icon
function slide03(pres) {
	const s = pres.addSlide();
	rect(s, 11.478, 5.813, 1.856, 1.687, { fill: { color: ORANGE } });
	imageBox(s, 6.313, 3.963, 6.09, 2.705);
	rect(s, 0, 3.963, 6.313, 2.705, { fill: { color: NAVY } });
	heading(s, 0.731, 1.504, 7.045, 'About Our Business', 28);
	eyebrow(s, 0.731, 1.1, 7.045, EYEBROW_ABOUT);
	body(s, 0.731, 2.278, 7.045, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat sed lectus vestibulum mattis ullamcorper. Nulla facilisi cras fermentum odio eu feugiat pretium nibh. Est sit amet iki');
	s.addShape('hexagon', { x: 5.644, y: 4.738, w: 1.338, h: 1.154, fill: { color: ORANGE }, line: { type: 'none' } });
	iconBlueprint(s, 5.948, 4.948, 0.732, 0.709, NAVY);
	body(s, 0.731, 4.983, 4.182, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat sed', MUTED);
	subhead(s, 0.731, 4.678, 4.182, SUB, WHITE, 'justify');
	imageBox(s, 8.507, 0.832, 4.826, 2.705);
}

// 4 — about our business + KPI strip
function slide04(pres) {
	const s = pres.addSlide();
	imageBox(s, 1.004, 3.958, 4.148, 2.538);
	rect(s, 4.424, 5.848, 8.532, 1.652, { fill: { color: NAVY } });
	[['198+', 4.999, 0.849], ['1423+', 7.015, 1.0], ['1214K', 9.031, 0.958], ['2523M', 11.046, 1.131]].forEach(([v, x, w]) => {
		text(s, v, { x, y: 6.304, w, h: 0.438, fontFace: SEMI, fontSize: 20, bold: true, color: ORANGE, wrap: false });
		text(s, 'Your Text Here', { x, y: 6.742, w: 1.334, h: 0.303, fontFace: LATO, fontSize: 12, color: MUTED, wrap: false });
	});
	heading(s, 5.792, 1.798, 6.524, 'About Our Business', 28);
	eyebrow(s, 5.792, 1.394, 6.538, EYEBROW_ABOUT);
	body(s, 5.792, 2.541, 6.524, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat sed lectus vestibulum mattis ullamcorper. Nulla facilisi cras fermentum odio eu feugiat pretium');
	rect(s, 12.956, 0, 0.377, 7.5, { fill: { color: ORANGE } });
	body(s, 8.694, 4.366, 3.621, 0.667,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est');
	subhead(s, 8.694, 4.063, 3.635, SUB);
	label(s, 10.204, 5.244, 1.737, 0.304, 'LEARN MORE', NAVY, 'right');
	arrow(s, 12.098, 5.388, NAVY);
	imageBox(s, 5.568, 3.947, 2.551, 1.902);
	imageBox(s, 1.004, 1.004, 4.148, 2.538);
}

// 5 — about our business, two feature rows on navy
function slide05(pres) {
	const s = pres.addSlide();
	rect(s, 0, 4.366, 8.964, 2.402, { fill: { color: NAVY } });
	body(s, 9.403, 1.525, 3.197, 0.97,
		'Blandit massa enim nec dui nunc mattis. Ac feugiat sed lectus vestibulum mattis ullamcorper. Nulla facilisi cras quis nisl');
	subhead(s, 9.403, 1.222, 3.197, SUB);
	rect(s, 9.403, 6.218, 3.197, 0.549, { fill: { color: ORANGE } });
	label(s, 9.403, 6.341, 3.197, 0.303, 'BUSINESS PROJECT', NAVY, 'center');
	heading(s, 0.47, 1.715, 4.828, 'About Our Business', 28);
	eyebrow(s, 0.47, 1.311, 4.858, EYEBROW_ABOUT);
	body(s, 0.47, 2.771, 4.828, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat sed lectus vestibulum');
	rule(s, 0, 2.515, 1.758, ORANGE);
	const feat = 'Consequat semper viverra nam libero justo. Amet iki et massa nisl suscipit adipiscing bibendum';
	iconDrill(s, 0.653, 5.013, 0.667, 0.604, ORANGE);
	body(s, 1.525, 5.234, 2.641, 0.97, feat, MUTED);
	subhead(s, 1.525, 4.929, 2.641, SUB, WHITE, 'justify');
	iconClipboard(s, 4.714, 5.013, 0.5, 0.667, ORANGE);
	body(s, 5.419, 5.234, 2.641, 0.97, feat, MUTED);
	subhead(s, 5.419, 4.929, 2.641, SUB, WHITE, 'justify');
	imageBox(s, 5.767, 0.733, 3.197, 3.197);
	imageBox(s, 9.403, 3.021, 3.197, 3.197);
}

// 6 — meet our team, four portraits
function slide06(pres) {
	const s = pres.addSlide();
	rect(s, 0, 7.162, 13.333, 0.337, { fill: { color: NAVY } });
	rect(s, 0, 3.628, 13.333, 1.267, { fill: { color: ORANGE } });
	[0.552, 3.747, 6.943, 10.138].forEach((x) => {
		text(s, 'Job Descriptions', { x, y: 6.099, w: 2.643, h: 0.303, fontFace: LATO, fontSize: 12, color: GRAY65, align: 'center' });
		text(s, 'Your Name Here', { x, y: 5.763, w: 2.643, h: 0.337, fontFace: SEMI, fontSize: 14, bold: true, color: GRAY_59, align: 'center' });
		socialRow(s, x + 0.819, 6.673);
	});
	heading(s, 2.377, 0.78, 8.58, 'Meet Our Team', 32, NAVY, 'center');
	eyebrow(s, 2.254, 0.376, 8.825, EYEBROW_KAYU, 18, 'center');
	body(s, 2.929, 1.632, 8.028, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat sed lectus vestibulum mattis ullamcorper. Nulla facilisi cras fermentum odio eu feugiat pretium nibh. Est sit amet facilisis magna etiam tempor orci',
		GRAY50, 'center');
	[0.593, 3.789, 6.984, 10.179].forEach((x) => imageBox(s, x, 3.158, 2.561, 2.207));
}

// 7 — talented team, skill bars
function slide07(pres) {
	const s = pres.addSlide();
	rect(s, 0.928, 1.083, 0.603, 2.538, { fill: { color: ORANGE } });
	rect(s, 5.49, 4.47, 7.844, 3.03, { fill: { color: NAVY } });
	heading(s, 6.293, 1.865, 6.238, 'Our Talented Team', 32);
	eyebrow(s, 6.293, 1.461, 6.113, EYEBROW_KAYU);
	text(s, 'Photography', { x: 6.293, y: 3.077, w: 6.238, h: 0.303, fontFace: LATO, fontSize: 12, color: GRAY50 });
	text(s, 'Daniel Edelmar', { x: 6.293, y: 2.741, w: 6.238, h: 0.337, fontFace: SEMI, fontSize: 14, bold: true, color: NAVY });
	body(s, 6.293, 3.38, 6.238, 0.675,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed un do eiusmod tempor incididunt ut labore et dolore magna aliqua urna cursus eget nunc et');
	// skill bars: [top of label, % label, filled width]
	[[4.841, '70%', 4.466], [5.705, '60%', 3.75], [6.584, '87%', 5.397]].forEach(([y, pct, filled]) => {
		text(s, pct, { x: 10.606, y, w: 1.925, h: 0.303, fontFace: SEMI, fontSize: 12, bold: true, color: WHITE, align: 'right' });
		text(s, 'Your Title Here', { x: 6.293, y, w: 2.273, h: 0.303, fontFace: SEMI, fontSize: 12, bold: true, color: WHITE });
		rule(s, 6.444, y + 0.392, 5.969, TRACK, 6.25);
		rule(s, 6.444, y + 0.392, filled, ORANGE, 6.25);
	});
	s.addText('MEET THE TEAM', {
		x: 1.027, y: 1.273, w: 0.404, h: 2.14, vert: 'vert270', valign: 'top',
		fontFace: POPPINS, fontSize: 12, bold: true, charSpacing: 1, color: NAVY, align: 'center',
	});
	imageBox(s, 1.531, 1.083, 3.959, 6.417);
}

// 8 — our services, 2x2 hexagon icons
function slide08(pres) {
	const s = pres.addSlide();
	const card = 'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer';
	[[7.103, 0.67, 6.144, 2.171], [10.534, 0.67, 9.575, 2.171], [7.103, 4.054, 6.144, 5.555], [10.534, 4.054, 9.575, 5.555]]
		.forEach(([hx, hy, tx, ty]) => {
			s.addShape('hexagon', { x: hx, y: hy, w: 1.338, h: 1.154, fill: { color: ORANGE }, line: { type: 'none' } });
			body(s, tx, ty + 0.305, 3.2, 0.97, card, GRAY50, 'center');
			subhead(s, tx, ty, 3.2, SUB, NAVY, 'center');
		});
	rect(s, 0.658, 0.67, 4.928, 3.092, { fill: { color: NAVY } });
	heading(s, 0.925, 1.514, 4.272, 'Our Services', 32, WHITE, 'center');
	eyebrow(s, 0.925, 1.143, 4.272, EYEBROW_KAYU, 16, 'center');
	body(s, 1.048, 2.318, 4.149, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer Blandit massa enim nec dui nunc mattis iki', MUTED, 'center');
	iconBulldozer(s, 7.409, 0.936, 0.727, 0.591, NAVY);
	iconCrane(s, 10.849, 0.892, 0.709, 0.709, NAVY);
	iconBoot(s, 7.399, 4.339, 0.748, 0.584, NAVY);
	iconPlug(s, 10.829, 4.257, 0.748, 0.748, NAVY);
	imageBox(s, 0.658, 4.054, 4.928, 3.446);
}

// 9 — our services, 2 white columns + 3 navy columns
function slide09(pres) {
	const s = pres.addSlide();
	rect(s, 0, 4.071, 13.333, 3.429, { fill: { color: NAVY } });
	const dark = 'Consequat semper viverra nam libero justo. Amet iki et nisl suscipit adipiscing bibendum est ultricies integer Blandit';
	[0.933, 5.067, 9.2].forEach((tx) => {
		body(s, tx, 5.86, 3.2, 0.97, dark, MUTED);
		subhead(s, tx, 5.555, 3.2, SUB, WHITE, 'justify');
	});
	iconWheelbarrow(s, 1.045, 4.637, 0.787, 0.689, ORANGE);
	iconHammer(s, 5.179, 4.588, 0.738, 0.787, ORANGE);
	iconScrewdriver(s, 9.312, 4.607, 0.748, 0.748, ORANGE);
	heading(s, 0.927, 0.846, 7.354, 'Our Services', 32, NAVY, 'justify');
	eyebrow(s, 0.927, 0.476, 7.354, EYEBROW_KAYU, 16, 'justify');
	const lightCol = 'Consequat semper viverra nam libero justo. Amet nisl animo ba iki nam suscipit adipiscing bibendum est ultricies blan';
	[0.933, 5.067].forEach((x) => {
		body(s, x, 2.459, 3.2, 0.97, lightCol);
		subhead(s, x, 2.156, 3.215, SUB);
	});
	rule(s, 1.045, 1.708, 0.787, ORANGE);
	rect(s, 0, 0, 0.308, 4.069, { fill: { color: ORANGE } });
	imageBox(s, 9.185, 0, 4.149, 4.069);
}

// 10 — our services, three navy cards
function slide10(pres) {
	const s = pres.addSlide();
	imageBox(s, 6.448, 0.711, 6.886, 4.126);
	[0.711, 4.101, 7.491].forEach((x) => rect(s, x, 4.123, 3.015, 3.377, { fill: { color: NAVY } }));
	iconVest(s, 1.886, 4.674, 0.664, 0.787, ORANGE);
	iconJackhammer(s, 5.215, 4.674, 0.787, 0.787, ORANGE);
	iconHelmet(s, 8.641, 4.674, 0.714, 0.787, ORANGE);
	[0.957, 4.347, 7.737].forEach((x) => {
		body(s, x, 5.98, 2.523, 0.97,
			'Consequat semper viverra nam libero justo. Amet iki et nisl suscipit adipiscing amet', MUTED, 'center');
		subhead(s, x, 5.675, 2.523, SUB, WHITE, 'center');
	});
	rect(s, 6.061, 0.709, 0.387, 3.412, { fill: { color: ORANGE } });
	heading(s, 0.727, 1.75, 4.623, 'Our Services', 32);
	eyebrow(s, 0.727, 1.377, 4.623, EYEBROW_KAYU, 16);
	body(s, 0.711, 2.603, 4.639, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat sed lectus');
}

// 11 — break slide
function slide11(pres) {
	const s = pres.addSlide();
	rect(s, 0, 0.831, 0.544, 6.669, { fill: { color: NAVY } });
	text(s, 'break slides', { x: 5.608, y: 2.907, w: 7.161, h: 1.01, fontFace: SEMI, fontSize: 54, bold: true, color: NAVY });
	text(s, 'It\u2019s Time For a Break', { x: 5.608, y: 3.916, w: 7.161, h: 0.438, fontFace: POPPINS, fontSize: 20, color: GRAY50 });
	rect(s, 5.044, 4.882, 8.289, 1.787, { fill: { color: ORANGE } });
	body(s, 5.608, 5.291, 7.161, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat sed lectus vestibulum mattis ullamcorper. Nulla facilisi cras fermentum odio eu feugiat pretium nibh. Est sit amet facilisis', NAVY);
	s.addText('CONSTRUCTION', {
		x: 0.07, y: 4.977, w: 0.404, h: 2.236, vert: 'vert270', valign: 'top',
		fontFace: POPPINS, fontSize: 12, bold: true, charSpacing: 1, color: ORANGE,
	});
	s.addShape('line', { x: 0.272, y: 2.769, w: 0, h: 2.521, line: { color: ORANGE, width: 1.5 } });
	imageBox(s, 0.544, 0.831, 4.5, 5.838);
}

// 12 — latest work, navy intro panel + four captions
function slide12(pres) {
	const s = pres.addSlide();
	rect(s, 6.667, 0, 0.451, 4.294, { fill: { color: ORANGE } });
	rect(s, 0, 0, 6.667, 4.294, { fill: { color: NAVY } });
	heading(s, 0.7, 1.184, 5.305, 'Our Latest Work', 32, WHITE);
	eyebrow(s, 0.7, 0.78, 5.305, EYEBROW_KAYU);
	body(s, 0.662, 2.141, 5.343, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis. Ac feugiat sed lectus vestibulum mattis ullamcorper iki', MUTED);
	const wide = 'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec';
	[0.629, 2.134].forEach((y) => {
		body(s, 7.578, y + 0.303, 5.294, 0.667, wide);
		subhead(s, 7.578, y, 5.294, SUB);
	});
	const narrow = 'Consequat semper viverra nam libero justo. Amet iki nisl suscipit adipiscing bibendum';
	[3.758, 5.565].forEach((y) => {
		body(s, 10.49, y + 0.303, 2.382, 0.97, narrow);
		subhead(s, 10.49, y, 2.382, SUB);
	});
	imageBox(s, 0.662, 3.75, 3.191, 3.75);
	imageBox(s, 4.515, 3.75, 3.191, 3.75);
	imageBox(s, 8.368, 3.75, 1.662, 1.28);
	imageBox(s, 8.368, 5.558, 1.662, 1.28);
}

// 13 — latest work, navy caption block over a photo
function slide13(pres) {
	const s = pres.addSlide();
	rect(s, 8.304, 2.471, 7.5, 2.559, { fill: { color: ORANGE }, rotate: 90 });
	imageBox(s, 4.162, 1.279, 7.892, 3.706);
	rect(s, 1.279, 3.279, 4.735, 1.706, { fill: { color: NAVY } });
	heading(s, 1.641, 4.023, 4.246, 'Our Latest Work', 32, WHITE);
	eyebrow(s, 1.641, 3.619, 4.246, EYEBROW_KAYU);
	[1.641, 5.887].forEach((x, i) => {
		body(s, x, 5.758, 3.368, 0.667,
			'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing iki bibendum');
		subhead(s, x, 5.455, 3.368, SUB);
		label(s, [2.889, 7.135][i], 6.711, 1.737, 0.304, 'LEARN MORE', NAVY, 'right');
		arrow(s, [4.783, 9.028][i], 6.855, NAVY);
	});
}

// 14 — latest work, numbered 01/02/03
function slide14(pres) {
	const s = pres.addSlide();
	imageBox(s, 4.99, 0.699, 3.676, 3.147);
	imageBox(s, 9.162, 0.699, 3.676, 3.147);
	rect(s, 0, 3.846, 4.99, 2.956, { fill: { color: NAVY } });
	const hexCopy = 'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum semper viverra';
	[[6.159, 6.25, 5.157, '02'], [10.331, 10.5, 9.328, '03']].forEach(([hx, nx, tx, num]) => {
		s.addShape('hexagon', { x: hx, y: 3.269, w: 1.338, h: 1.154, fill: { color: ORANGE }, line: { type: 'none' } });
		text(s, num, { x: nx, y: 3.492, w: num === '02' ? 1.157 : 1.0, h: 0.707, fontFace: SEMI, fontSize: 36, bold: true, color: NAVY, align: 'center' });
		body(s, tx, 5.346, 3.343, 0.97, hexCopy, GRAY50, 'center');
		subhead(s, tx, 5.043, 3.343, SUB, BLACK, 'center');
	});
	subhead(s, 0.479, 5.192, 4.016, SUB, WHITE, 'justify');
	body(s, 0.479, 5.495, 4.016, 0.667,
		'PLACEHOLDER', MUTED);
	text(s, '01', { x: 0.479, y: 4.485, w: 0.98, h: 0.707, fontFace: SEMI, fontSize: 36, bold: true, color: ORANGE });
	heading(s, 0.479, 1.524, 4.016, 'Our Latest Work', 32);
	eyebrow(s, 0.479, 1.12, 4.016, EYEBROW_KAYU);
	body(s, 0.479, 2.342, 4.016, 0.97,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec dui nunc mattis ac est');
}

// 15 — latest work, two stats on a navy band
function slide15(pres) {
	const s = pres.addSlide();
	rect(s, 0.662, 1.39, 10.853, 4.721, { fill: { color: NAVY } });
	rect(s, 11.515, 1.39, 0.662, 4.721, { fill: { color: ORANGE } });
	heading(s, 5.534, 2.208, 5.337, 'Our Latest Work', 32, WHITE);
	eyebrow(s, 5.534, 1.804, 5.337, EYEBROW_KAYU);
	body(s, 5.534, 3.016, 5.343, 0.667,
		'Consequat semper viverra nam libero justo. Amet nisl suscipit adipiscing bibendum est ultricies integer quis. Blandit massa enim nec', MUTED);
	[[5.534, 5.559, '225+'], [8.471, 8.496, '233+']].forEach(([x, nx, num]) => {
		subhead(s, x, 4.565, 2.4, SUB, WHITE, 'justify');
		body(s, x, 4.868, 2.407, 0.667, 'Bibendum es ultricies integer quis blandit massa enim nec', MUTED);
		text(s, num, { x: nx, y: 3.987, w: 2.376, h: 0.572, fontFace: SEMI, fontSize: 28, bold: true, color: ORANGE });
	});
	imageBox(s, 1.324, 0.787, 3.574, 5.926);
}

// 16 — pricing plans
function slide16(pres) {
	const s = pres.addSlide();
	rect(s, 0, 0, 1.223, 7.5, { fill: { color: ORANGE } });
	rect(s, 4.575, 1.117, 3.527, 4.561, { fill: { type: 'none' }, line: { color: ORANGE, width: 2 } });
	rect(s, 0.616, 1.117, 3.527, 5.265, { fill: { color: NAVY }, line: { color: NAVY, width: 2 } });

	const FEATURES = ['Eiusmod tempor incididunt uta', 'Tempor incididunt ut labore', 'Quis nostrud exercitation', 'Commodo consequat', 'Consequat duis'];
	function plan(x, yTitle, title, price, color, featureTop, features) {
		text(s, title, { x, y: yTitle, w: 2.665, h: 0.572, fontFace: SEMI, fontSize: 28, bold: true, color, align: 'center' });
		text(s, 'Start From', { x, y: yTitle + 0.718, w: 2.665, h: 0.337, fontFace: SEMI, fontSize: 14, bold: true, color, align: 'center' });
		text(s, price, { x: x + 0.011, y: yTitle + 1.138, w: 2.645, h: 0.64, fontFace: SEMI, fontSize: 32, bold: true, color, align: 'center' });
		features.forEach((f, i) => {
			s.addText(f, {
				x: x + 0.011, y: featureTop + i * 0.463, w: 2.645, h: 0.303, valign: 'top',
				fontFace: LATO, fontSize: 12, color: color === WHITE ? MUTED : GRAY50,
				bullet: { characterCode: '2713', indent: 13.5 },
			});
		});
	}
	plan(1.047, 1.671, 'Enterprise', '$399', WHITE, 3.675, FEATURES);
	plan(5.006, 1.549, 'Standard', '$199', GRAY_40, 3.552, FEATURES.slice(0, 4));

	heading(s, 8.72, 2.108, 3.995, 'Pricing Plans', 32, GRAY_40);
	eyebrow(s, 8.72, 1.704, 3.995, 'Spavin Company');
	body(s, 8.72, 3.317, 3.995, 0.97,
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua massa sapien faucibus et sed do');
	rule(s, 8.833, 3.002, 0.824, ORANGE, 2);
	rect(s, 8.794, 4.623, 1.754, 0.469, { fill: { color: NAVY } });
	label(s, 8.794, 4.735, 1.754, 0.303, 'SELECT PLAN', WHITE, 'center', POPPINS, 11);
}

// 17 — contact us
function slide17(pres) {
	const s = pres.addSlide();
	imageBox(s, 0, 1.791, 13.333, 3.0);
	const CONTACTS = [
		[1.313, 0.766, 'Location', '1234 Queens Bayside Point, \nCA 12345 United States ', iconPin, 1.726, 4.477, 0.512, 0.63],
		[4.455, 3.908, 'Office Hours', 'Monday-Friday \n09.00-17.00 ', iconClock, 4.809, 4.477, 0.63, 0.63],
		[7.597, 7.05, 'Get In Touch', '(+62) 123 5678 000\n(+62) 123 4678 900 ', iconPhone, 7.951, 4.477, 0.63, 0.63],
		[10.738, 10.191, 'More Information', 'www.volivatravel.com  contact@voliva.com', iconMail, 11.093, 4.561, 0.63, 0.462],
	];
	CONTACTS.forEach(([hx, tx, title, lines, glyph, ix, iy, iw, ih]) => {
		s.addShape('hexagon', { x: hx, y: 4.215, w: 1.338, h: 1.154, fill: { color: ORANGE }, line: { type: 'none' } });
		glyph(s, ix, iy, iw, ih, NAVY);
		text(s, lines, {
			x: tx, y: 6.082, w: 2.376, h: 0.675, fontFace: OPENS, fontSize: 12,
			color: GRAY65, align: 'center', lineSpacingMultiple: 1.5,
		});
		subhead(s, tx, 5.777, 2.376, title, NAVY, 'center');
	});
	rect(s, 0, 7.162, 13.333, 0.337, { fill: { color: NAVY } });
	heading(s, 3.612, 0.724, 6.109, 'Contact Us', 36, NAVY, 'center');
	eyebrow(s, 3.603, 0.348, 6.118, EYEBROW_KAYU, 16, 'center');
}

// 18 — thank you
function slide18(pres) {
	const s = pres.addSlide();
	imageBox(s, 0, 0, 13.333, 4.093);
	rect(s, 0, 7.0, 11.794, 0.5, { fill: { color: ORANGE } });
	rect(s, 1.113, 2.77, 6.324, 2.647, { fill: { color: NAVY } });
	text(s, 'thank you', { x: 1.669, y: 3.386, w: 4.819, h: 1.01, fontFace: SEMI, fontSize: 54, bold: true, color: ORANGE });
	text(s, 'Construction Presentation', { x: 1.669, y: 4.396, w: 4.819, h: 0.404, fontFace: POPPINS, color: MUTED });
	text(s, '2023', { x: 12.231, y: 7.065, w: 0.738, h: 0.37, fontFace: SEMI, fontSize: 16, bold: true, color: GRAY50, wrap: false });
}

/* --------------------------------------------------------------------- main */
function build() {
	const pres = new PptxGenJS();
	pres.defineLayout({ name: 'WIDE', width: 13.3333333, height: 7.5 });
	pres.layout = 'WIDE';
	pres.title = 'kayuadam — Construction Presentation';

	[slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09,
		slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18]
		.forEach((fn) => fn(pres));

	return pres.writeFile({ fileName: path.join(__dirname, '1aeca05f-1acf-49eb-816b-e56d2defb74f_grok_final.pptx') });
}

build().then((f) => console.log('wrote', f)).catch((e) => { console.error(e); process.exit(1); });
