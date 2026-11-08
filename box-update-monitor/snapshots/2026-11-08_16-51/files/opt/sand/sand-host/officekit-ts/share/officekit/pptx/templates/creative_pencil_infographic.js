/**
 * Creative Pencil Solutions — 25-slide infographic deck, 13.333in x 7.5in.
 *
 * The whole deck is plain pptxgenjs calls: outline tables, per-slide builder
 * functions and two small helpers (`draw` for shapes, `tx` for text).
 * Raster art in the source deck (icon glyphs, photographs) is reproduced as
 * flat colour placeholder blocks in the same position and size.
 *
 *   node <thisfile>.js   ->  writes the .pptx beside this script
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ── palette ─────────────────────────────────────────────────────────────────
// Theme accents, plus the tint/shade variants and neutrals the artwork uses.
const BLUE = '3E3DF2', PINK = 'FD5EA8', TEAL = '00DCBA', GOLD = 'FDBF21', PURPLE = '4E3F9C', SLATE = '415A78';
const BLUE_D = '0F0ED5', BLUE_XD = '0A0A8E', BLUE_L = '8B8BF7', BLUE_P = 'D8D8FC';
const PINK_D = 'FC087A', PINK_XD = 'AB0251', PINK_L = 'FE9ECB', PINK_P = 'FEBFDC', PINK_W = 'FFDFEE';
const TEAL_D = '00A58B', TEAL_XD = '006E5D', TEAL_L = '51FFE4', TEAL_P = '8BFFED', TEAL_W = 'C5FFF6';
const GOLD_D = 'D59902', GOLD_XD = '8E6601', GOLD_L = 'FED97A', GOLD_W = 'FFF2D3', GOLD_O = 'FFB14A';
const SLATE_D = '31445A', SLATE_L = '819BBA', SLATE_T = '333F50';
const PURPLE_D = '3B2F75', PURPLE_L = '8F83CD';
const WHITE = 'FFFFFF', SNOW = 'F2F2F2', GREY = 'E7E6E6', GREY2 = 'D9D9D9', GREY3 = 'D0CECE';
const GREY4 = 'BFBFBF', GREY5 = 'AFABAB', INK_L = '595959', INK = '404040', CHAR = '262626';
const CHAR2 = '231F20', CHAR3 = '404041', CHAR4 = '58595B', BLACK = '000000';
const SKIN = 'FEDCB8', SKIN_L = 'FEE9D2', SKIN_D = 'FECF90', SKIN_M = 'FFD29E';
// dominant colours of the raster icons / photos that the placeholders replace
const IC_BLUE = '3030F0', IC_PINK = 'F048A8', IC_TEAL = '00D8A8', IC_GOLD = 'F0A818';
const IC_SLATE = '304878', IC_PURPLE = '483090', IC_WHITE = 'F0F0F0';

const HEAD = 'Sora SemiBold';   // theme major font
const BODY = 'Open Sans';       // theme minor font

// shape ids usable in the first slot of a `draw` row
const RECT = 'rect', LINE = 'line', IMG = '#image';

// ── slide furniture inherited from the master ───────────────────────────────
function chrome(s, pageNo, onDark) {
  const c = onDark ? WHITE : BLUE;
  s.addText('Creative Pencil', { x: 0.4, y: 0.44, w: 1.6, h: 0.31, fontSize: 12, fontFace: HEAD, color: c, valign: 'middle' });
  s.addText(String(pageNo), { x: 12.47, y: 6.74, w: 0.44, h: 0.28, fontSize: 10, fontFace: HEAD, color: c, align: 'center', valign: 'middle' });
}

/**
 * Slide 1's 45-degree two-stop wash. OOXML linear gradients are not available
 * for slide backgrounds through pptxgenjs, so it is painted as a stack of
 * narrow bands drawn perpendicular to the gradient axis.
 */
function diagonalGradient(s, from, to) {
  s.background = { color: to };
  const BANDS = 60, DIAG = (13.333 + 7.5) * Math.SQRT1_2, STOP = 0.75;
  const dx = Math.SQRT1_2, dy = Math.SQRT1_2;      // axis points down-right
  for (let i = 0; i < BANDS; i++) {
    const t = (i + 0.5) / BANDS;                   // 0 at top-left corner, 1 at bottom-right
    const cx = 6.667 + dx * (t - 0.5) * DIAG;
    const cy = 3.75 + dy * (t - 0.5) * DIAG;
    const bw = 22, bh = DIAG / BANDS + 0.06;
    s.addShape('rect', {
      x: cx - bw / 2, y: cy - bh / 2, w: bw, h: bh, rotate: -45,
      fill: { color: mix(from, to, Math.min(1, t / STOP)) }, line: { type: 'none' },
    });
  }
}

function mix(a, b, t) {
  const ch = (h, i) => parseInt(h.substr(i * 2, 2), 16);
  const v = i => Math.round(ch(a, i) + (ch(b, i) - ch(a, i)) * t).toString(16).padStart(2, '0');
  return (v(0) + v(1) + v(2)).toUpperCase();
}

/**
 * Draw a table of shapes. Each row is
 *     [shape, x, y, w, h, fill, extras?]
 * where `shape` is an outline table (percent polygon), an array of outline
 * tables (2nd and later ones cut holes), RECT, LINE, IMG, or any pptxgenjs
 * preset name. Sizes are inches. `extras` may carry:
 *     r   rotation in degrees          fv  flip vertically
 *     tr  fill transparency 0-100      sd  drop shadow
 *     lc  line colour   lw  line width pt   ld  dashed line
 *     rr  corner radius (roundRect)
 *     a1,a2 start/sweep angle (arc, blockArc)   a3 arc thickness
 */
function draw(s, rows) {
  rows.forEach(([shape, x, y, w, h, fill, ex]) => {
    ex = ex || {};
    const o = { x, y, w, h };
    if (ex.r) o.rotate = ex.r;
    if (ex.fv) o.flipV = true;
    o.fill = fill ? { color: fill, transparency: ex.tr || 0 } : { type: 'none' };
    o.line = ex.lc ? { color: ex.lc, width: ex.lw || 1, dashType: ex.ld ? 'dash' : 'solid' } : { type: 'none' };
    if (ex.sd) o.shadow = { type: 'outer', color: BLACK, opacity: 0.18, blur: 12, offset: 2, angle: 90 };
    if (ex.rr !== undefined) o.rectRadius = ex.rr;
    if (ex.a1 !== undefined) o.angleRange = [ex.a1, ex.a2];
    if (ex.a3 !== undefined) o.arcThicknessRatio = ex.a3;

    if (shape === IMG) {                       // stand-in for a raster asset
      o.fill = { color: fill || GREY };
      if (Math.max(w, h) <= 0.9) { s.addShape('ellipse', o); return; }   // small icon glyph
      s.addShape(RECT, o);
      s.addText('[image]', { x, y, w, h, align: 'center', valign: 'middle', fontSize: 11, color: WHITE, fontFace: BODY });
      return;
    }
    if (typeof shape === 'string') { s.addShape(shape, o); return; }

    const rings = Array.isArray(shape[0]) ? shape : [shape];
    o.points = [];
    rings.forEach(ring => {
      for (let i = 0; i < ring.length; i += 2) {
        o.points.push({ x: w * ring[i] / 100, y: h * ring[i + 1] / 100, moveTo: i === 0 });
      }
      o.points.push({ close: true });
    });
    s.addShape('custGeom', o);
  });
}

const BR = { _br: true };   // hard line break inside a text box

/**
 * Text box. Options: x,y,w,h inches; sz pt; c colour; f font face;
 * a align; ls line-spacing multiple; sb space-before pt; v vertical centre;
 * nw no wrap; af shrink box to text; f2 box fill; sd drop shadow.
 * Runs are plain strings, ['text', {overrides}] pairs, or BR.
 */
function tx(s, o, runs) {
  const body = [];
  runs.forEach(r => {
    if (r === BR) { if (body.length) body[body.length - 1].options.breakLine = true; return; }
    const [text, m] = Array.isArray(r) ? r : [r, {}];
    body.push({
      text,
      options: {
        fontSize: m.sz || o.sz, color: m.c || o.c, bold: !!m.b, italic: !!m.i,
        fontFace: m.f || o.f || BODY, breakLine: false,
      },
    });
  });
  s.addText(body, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    align: o.a || 'left', valign: o.v ? 'middle' : 'top',
    fontSize: o.sz, color: o.c, fontFace: o.f || BODY,
    lineSpacingMultiple: o.ls || 1.0, paraSpaceBefore: o.sb || 0,
    wrap: !o.nw, autoFit: !!o.af, margin: [7.2, 7.2, 3.6, 3.6],   // OOXML default text insets
    fill: o.f2 ? { color: o.f2 } : undefined,
    shadow: o.sd ? { type: 'outer', color: BLACK, opacity: 0.18, blur: 12, offset: 2, angle: 90 } : undefined,
    line: { type: 'none' },
  });
}

// ── artwork outlines ────────────────────────────────────────────────────────
// Each outline is a closed polygon given in PERCENT of its own shape box:
// [x0,y0, x1,y1, ...]. Names are P<slide><letter> after first appearance.
const P1A = [47,100,13,99,0,95,27,43,47,1,47,0,100,0,100,9,100,95,83,99];
const P1B = [47,100,14,99,0,95,1,73,4,46,9,24,15,9,20,0,79,0,79,3,78,18,82,42,90,70,100,95,83,99];
const P1C = [47,100,14,98,1,93,1,1,1,0,78,0,81,16,100,93,83,98];
const P1D = [1,31,50,100,53,98,76,66,99,31,100,28,85,18,78,3,75,0,65,19,58,24,50,26,42,24,36,19,31,12,27,1,25,0,22,3,16,18,0,28,1,31,33,66,50,64,67,66,50,90,33,66];
const P1E = [47,100,13,98,0,93,9,14,10,0,94,0,95,16,100,93,83,98];
const P4A = [0,25,65,0,100,32,100,68,65,100,0,75];
const P4B = [93,100,100,0,0,0,0,100];
const P4C = [100,100,93,0,0,0,0,100];
const P4D = [84,100,0,50,84,0,100,0,100,100];
const P4E = [100,0,9,0,3,15,1,34,1,65,3,85,9,100,100,100];
const P4F = [100,56,94,0,0,50,96,100];
const P4G = [100,0,17,0,13,1,10,4,5,15,1,30,0,50,1,69,5,85,10,96,13,99,17,100,100,100];
const P4H = [0,11,100,0,100,85,0,100];
const P5A = [100,100,50,0,0,100];
const P5B = [100,100,2,0,0,100];
const P5C = [98,100,100,0,99,0,0,100];
const P5D = [100,100,0,100,48,1,50,0];
const P5E = [96,100,0,100,96,1,100,0];
const P5F = [0,0,0,94,15,98,34,99,66,99,85,98,100,94,100,0];
const P5G = [0,0,0,75,15,91,34,98,66,98,85,91,100,75,100,0];
const P5H = [99,100,100,0,0,100];
const P5I = [100,100,0,100,100,1];
const P5J = [0,0,0,94,4,96,15,98,34,100,66,100,85,98,96,96,100,94,100,0];
const P6A = [77,13,67,14,18,0,0,4,74,98,85,100,100,97];
const P6B = [81,0,33,14,23,13,0,96,15,100,26,98,100,5];
const P6C = [0,13,31,97,50,100,69,96,100,13,50,0];
const P6D = [36,0,0,100,100,4,74,7];
const P6E = [64,0,26,6,0,3,100,100];
const P6F = [50,7,0,1,50,100,100,0];
const P6G = [0,0,100,0,100,97,92,96,71,100,67,100,50,96,33,100,29,100,8,96,0,97];
const P6H = [0,0,100,0,100,97,76,96,13,100,0,100];
const P6I = [0,0,100,0,100,100,50,97,0,100];
const P7A = [1,23,6,18,13,8,18,8,29,17,45,22,58,17,63,13,70,12,83,21,88,23,93,22,100,12,97,21,68,88,64,95,62,94,54,94,53,97,48,97,36,93];
const P7B = [98,0,25,0,23,35,11,34,3,39,0,50,3,61,11,66,23,65,25,100,50,100,49,84,61,71,70,75,74,84,73,100,98,100];
const P7C = [84,34,100,35,97,0,65,0,50,29,35,0,0,2,0,97,35,100,50,71,65,100,97,100,100,65,84,66,70,50];
const P7D = [95,16,91,27,65,88,53,99,42,86,13,25,12,19,16,21,69,10];
const P7E = [89,25,77,26,75,0,50,0,51,11,38,21,26,11,27,0,2,0,2,83,7,86,22,74,58,90,77,88,77,47,89,48,100,36];
const P7F = [50,0,49,12,61,23,74,12,71,1,100,2,100,75,73,77,70,97,61,100,50,77,25,77,23,50,0,39,3,30,23,27,25,0];
const P7G = [53,27,38,52,21,64,11,72,11,83,18,91,30,90,63,64,84,33,87,23,84,17];
const P7H = [68,30,63,19,51,10,20,0,9,4,7,15,30,16,29,20,4,27,8,36,27,33,27,38,13,39,3,45,0,54,4,57,27,54,28,62,12,76,11,100,22,81,30,77,58,82,65,76,100,96,100,48];
const P7I = [100,28,11,0,0,70,100,100];
const P7J = [100,36,7,0,0,62,100,100];
const P7K = [80,7,88,56,100,79,29,77,8,61,0,43,24,15];
const P7L = [77,0,100,73,47,82,21,63,18,47,25,28];
const P7M = [100,74,60,0,20,30,18,62,45,82];
const P7N = [100,68,57,0,20,35,20,68,49,85];
const P11A = [100,90,54,0,44,0,0,94,3,100,95,100];
const P11B = [100,100,0,100,0,2,19,1,49,0,80,1,100,2];
const P11C = [0,92,100,93,59,0,35,0];
const P11D = [46,0,11,2,0,2,0,86,3,86,27,89,42,100,54,88,73,84,100,93];
const P11E = [49,0,48,0,18,1,0,2,0,100,2,100,100,100,100,2,90,1];
const P11F = [27,0,16,0,0,0,0,100,24,100,100,97];
const P11G = [9,0,7,0,0,0,0,100,100,100,98,100,98,2,61,0];
const P12A = [23,16,16,18,7,37,7,63,16,81,23,84,29,81,38,63,38,37,29,18,23,16,27,0,44,4,66,22,86,53,100,88,79,80,65,82,30,99,18,98,6,83,0,51,7,16,24,0,27,1];
const P12B = [31,7,16,13,11,19,11,33,16,39,31,45,47,39,53,26,51,19,47,13,31,7,33,0,51,5,70,21,87,44,97,73,100,100,75,77,7,44,0,29,5,12,22,1,33,0];
const P12C = [50,6,37,7,19,16,16,22,26,32,50,37,63,36,81,28,84,22,81,16,63,7,50,6,49,0,87,8,95,13,100,24,95,42,65,75,54,100,52,88,46,76,2,36,0,23,2,17,17,6,49,0];
const P12D = [74,16,68,18,62,26,59,37,62,74,74,84,81,81,90,63,90,37,81,18,74,16,73,0,86,5,93,16,100,51,94,83,82,98,70,99,35,82,21,80,0,88,14,53,34,22,56,4,73,0];
const P12E = [67,8,47,20,47,34,59,44,76,44,83,40,88,34,88,20,83,14,67,8,67,0,78,1,95,12,99,21,98,37,93,44,25,77,0,100,3,73,13,44,30,21,49,5,67,0];
const P12F = [18,99,50,99,81,100,94,95,99,83,100,65,95,54,54,2,52,0,49,2,5,55,1,65,1,83,5,95];
const P12G = [100,88,52,0,0,89,48,100];
const P12H = [0,12,7,38,41,88,65,88,69,75,93,38,100,0,85,62,65,100,41,100,30,88,13,62];
const P12I = [37,0,100,4,95,96,95,100,0,100,5,3];
const P12J = [75,0,99,2,95,94,95,100,0,100,1,90,6,4];
const P12K = [52,0,100,3,96,97,96,100,0,100,0,90,4,2];
const P13A = [0,0,95,0,100,0,100,97,92,96,71,100,67,100,50,96,33,100,29,100,8,96,0,97];
const P14A = [0,0,2,0,95,0,100,0,100,97,92,96,71,100,67,100,50,96,33,100,29,100,8,96,0,97];
const P14B = [100,0,0,0,50,100];
const P14C = [39,100,100,100,61,0,0,0];
const P18A = [89,100,0,99,11,0,100,0];
const P18B = [25,12,47,0,75,12,90,1,100,6,53,100,0,6,15,3];
const P18C = [0,9,53,100,100,9,45,0];
const P18D = [100,81,18,81,14,78,10,72,8,62,7,50,8,38,10,28,14,22,18,19,18,0,15,1,11,4,5,15,1,31,0,51,2,70,5,85,11,96,15,99,18,100,100,100];
const P18E = [100,72,14,72,12,70,10,65,8,50,10,34,12,30,14,28,14,0,11,1,9,4,4,15,1,31,0,50,1,70,4,86,9,96,12,99,14,100,100,100];
const P18F = [67,13,76,16,84,24,89,36,91,50,89,64,84,76,76,84,67,87,67,100,74,99,80,96,90,86,97,70,100,51,98,31,91,15,81,4,74,1,68,0,67,0,0,0,0,13];
const P18G = [0,19,74,19,80,22,85,28,89,38,90,50,89,62,85,72,80,78,74,81,74,100,79,99,84,96,92,85,98,69,100,49,98,30,92,15,84,4,79,1,74,0,0,0];
const P18H = [0,29,82,29,85,30,88,35,90,50,88,66,85,70,82,72,82,100,89,96,95,85,99,70,100,50,99,31,95,15,89,4,82,0,0,0];
const P18I = [0,10,62,10,68,11,74,13,84,22,90,34,92,50,90,65,84,78,74,86,69,89,62,90,62,100,70,99,77,96,83,92,89,85,97,70,100,50,97,31,89,15,83,9,77,4,70,1,62,0,0,0];
const P18J = [21,87,18,86,15,84,10,76,7,64,6,50,7,36,10,24,15,16,18,14,21,13,21,0,17,1,13,4,6,15,2,31,0,50,2,69,6,85,13,96,17,99,21,100,100,100,100,87];
const P18K = [100,90,23,90,19,89,16,86,10,78,6,65,5,50,6,35,10,22,16,13,19,11,23,10,23,0,18,1,14,4,10,9,7,15,2,31,0,50,2,69,7,85,10,91,14,96,18,99,23,100,100,100];
const P19A = [85,87,92,56,87,29,85,35,67,9,38,2,0,16,8,22,20,62,75,90,82,100,68,45,41,21,57,29,72,46];
const P19B = [5,55,16,27,47,10,97,5,91,12,94,48,38,100,34,86,39,48,61,20,29,49,23,90,7,73];
const P19C = [4,65,13,35,30,17,28,23,54,15,78,31,100,72,91,71,63,94,51,94,24,80,29,89,10,76,0,78,36,45,54,41,67,45,52,39,33,41];
const P19D = [4,65,13,35,30,17,28,23,54,15,78,31,100,72,91,71,63,94,51,94,24,80,29,89,10,76,0,78,36,45,67,45,52,39,33,41];
const P19E = [57,8,35,21,17,42,6,69,3,100,0,92,4,63,15,38,31,17,53,5,78,3,100,12,79,6];
const P19F = [57,99,14,100,14,40,7,9,12,2,24,0,55,3,73,10,77,20,67,35,59,60];
const P19G = [70,100,97,98,97,4,70,1,36,1,73,2,80,7,79,25,66,59,47,98];
const P19H = [31,100,61,97,52,68,49,40,55,22,65,14,78,7,46,14,31,20,16,37,6,66,4,97];
const P19I = [45,100,85,97,70,67,65,36,68,22,72,16,85,7,88,4,83,3,45,5,31,10,17,28,10,61,10,97];
const P19J = [1,31,50,100,76,66,100,28,85,18,79,4,75,1,65,19,50,26,35,19,28,4,23,1,15,18,0,28,33,66,50,64,67,66,50,90,33,66];
const P19K = [64,56,76,65,86,76,91,85,91,93,55,92,43,88,37,83,25,69,17,47,15,22,20,0,23,10,30,18,64,27,77,36,93,57,100,84,85,66];
const P19L = [4,65,13,35,30,17,28,23,54,15,78,31,100,72,91,71,73,90,63,94,51,94,24,80,29,89,10,76,0,78,36,45,54,41,67,45,52,39,33,41];
const P19M = [87,1,90,43,100,98,47,100,0,98,17,42,27,0];
const P19N = [85,87,92,56,87,29,85,35,67,9,38,2,0,16,8,22,20,62,52,85,75,90,82,100,68,45,41,21,57,29,72,46];
const P19O = [29,82,14,57,23,30,58,0,60,9,93,34,90,76,97,69,94,100,50,61,42,29,43,67,73,100,47,96];
const P19P = [10,42,27,14,56,3,80,3,100,10,92,17,83,57,53,83,23,100,39,43,65,17,30,41,13,85,5,63,7,37];
const P19Q = [5,55,16,27,47,10,97,5,91,12,94,48,54,80,38,100,34,86,39,48,61,20,29,49,23,90,7,73];
const P19R = [14,62,16,34,41,14,88,1,84,10,96,43,67,79,78,76,59,90,55,100,43,50,57,21,34,53,38,92,19,79,8,59];
const P19S = [85,87,92,56,87,29,85,35,67,9,38,2,0,16,8,22,20,62,82,100,68,45,41,21,57,29,72,46];
const P19T = [4,65,13,35,30,17,28,23,54,15,78,31,100,72,91,71,63,94,29,89,10,76,0,78,36,45,67,45,52,39,33,41];
const P19U = [100,0,94,94,58,78,90,58,90,11,42,11,36,94,6,94,32,0,100,0,10,78,79,89];
const P19V = [0,45,45,90,31,15,50,0,70,20,55,90,90,55,75,45,100,50,50,100,0,45,60,20];
const P19W = [55,90,73,40,27,40,55,90,0,65,10,26,90,26,90,94,5,86,50,0,70,20,30,20];
const P19X = [39,89,22,21,33,40,61,11,90,50,50,100,15,85,0,50,50,0,100,50,50,100,44,40];
const P20A = [100,100,33,0,0,50,33,100];
const P20B = [0,100,67,0,100,50,67,100];
const P21A = [100,2,100,76,71,78,73,89,70,97,61,100,53,97,49,89,52,78,25,78,23,50,11,51,3,48,0,39,3,30,11,27,23,28,25,0,50,0,52,2,49,12,53,20,62,23,70,20,74,12,73,0];
const P21B = [98,0,25,0,23,2,23,35,21,38,11,34,7,36,0,50,7,64,11,66,21,62,23,65,23,98,25,100,50,100,52,98,49,84,53,75,62,71,70,75,74,84,71,98,73,100,98,100,100,98,100,2];
const P21C = [98,18,73,18,71,17,74,9,70,3,62,0,53,3,49,9,52,17,25,18,23,41,3,43,1,54,11,60,23,59,25,82,50,82,52,83,49,91,53,97,62,100,70,97,74,91,71,83,73,82,100,80];
const P22A = [44,0,84,16,98,29,97,48,76,86,60,98,39,98,0,82];
const P22B = [100,50,85,85,50,100,15,85,0,50,15,15,50,0,85,15];
const P22C = [100,51,96,31,85,15,70,4,50,0,30,4,15,15,4,31,0,51,7,77,26,96,27,92,10,75,5,64,5,42,12,25,24,12,41,5,59,5,76,12,93,33,96,59,87,80,63,96,64,100,90,82];
const P22D = [100,51,96,31,85,15,69,4,50,0,30,4,15,15,4,31,0,51,7,77,26,96,28,91,14,79,6,63,5,46,11,28,18,19,38,7,59,6,75,13,92,33,94,59,86,79,63,95,64,100,90,82];
const P22E = [100,51,96,31,85,15,69,4,50,0,30,4,15,15,4,31,0,51,7,77,25,96,27,91,14,79,6,63,5,45,11,28,27,11,38,6,59,6,76,12,92,33,96,51,86,79,64,96,64,100,90,82];
const P22F = [96,100,12,86,0,47,68,0,79,13,76,31,92,47,87,69,99,83];
const P22G = [100,100,81,52,77,0,14,52,0,73,21,84];
const P23A = [90,82,100,0,0,0,0,100,89,100];
const P23B = [0,86,100,100,100,14,0,0];
const P23C = [100,70,93,100,4,100,0,70,24,0,76,0];
const P23D = [0,38,16,0,100,70,74,100];
const P23E = [5,100,0,0,100,0,91,100];
const P23F = [100,38,76,100,27,100,0,38,24,0,76,0];
const P23G = [48,0,100,100,0,18];
const P23H = [52,0,0,100,100,18];
const P23I = [100,29,81,0,23,0,0,22,52,100];

// ── body copy reused across slides ──────────────────────────────────────────
const T1 = 'Lorem ipsum dolor sit amet, adipiscing elit. ';
const T2 = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ';
const T3 = ' away, behind the word mountains, far from the';
const T4 = ' away, behind the word mountains, far from.';
const T5 = ' away, behind the word mountains, far from the countries ';
const T6 = ' away, behind the word mountains, far from';

// Slide 1 — Cover — gradient wash behind an oversized pencil outline
function slide01(s) {
  diagonalGradient(s, BLUE, BLUE_XD);
  chrome(s, 1, true);
  draw(s, [
    [P1A, 11.79,4.91,.27,2.59, BLUE, {r:180}],
    [P1B, 8.03,4.91,.27,2.59, BLUE, {r:180}],
    [P1C, 8.89,5.68,.27,1.82, BLUE, {r:180}],
    [P1D, 8.03,2.59,4.04,3.4, BLUE, {r:180}],
    [P1E, 10.93,5.68,.27,1.82, BLUE, {r:180}],
  ]);
  tx(s, {x:.94, y:4.17, w:9.97, h:1.01, sz:54, c:WHITE, f:HEAD, nw:1, af:1}, ['Creative Pencil Solutions']);
  tx(s, {x:.99, y:5.13, w:3.8, h:.37, sz:16, c:WHITE, nw:1, af:1}, ['Infographic Presentation Template']);
}

// Slide 2 — Two photo panels above a headline, rule and stat callout
function slide02(s) {
  chrome(s, 2);
  tx(s, {x:1.26, y:3.81, w:11.01, h:.74, sz:38, c:BLUE, f:HEAD, nw:1, af:1}, [['Pencil ', {c:BLACK}], 'Solutions for Maximum', [' Influence', {c:BLACK}]]);
  draw(s, [
    [LINE, 1.39,4.83,10.63,0, null, {lc:SNOW}],
  ]);
  tx(s, {x:10.44, y:5.24, w:1.27, h:.51, sz:24, c:BLUE, f:HEAD, nw:1, af:1}, ['$4.5M']);
  tx(s, {x:10.26, y:5.75, w:1.72, h:.52, sz:12, c:BLUE, f:HEAD, a:'center', v:1, f2:WHITE, sd:1}, ['Learn More']);
  tx(s, {x:1.29, y:5.22, w:8.34, h:1.13, sz:12, c:INK, ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ']);
}

// Slide 3 — Side photo, headline, doughnut KPI card
function slide03(s) {
  chrome(s, 3);
  tx(s, {x:5.54, y:1.28, w:6.31, h:2.12, sz:40, c:BLACK, f:HEAD, af:1}, ['The Art of ', ['Persuasion Pencil Solutions', {c:BLUE}], ' for Maximum Influence']);
  tx(s, {x:5.54, y:3.59, w:6.25, h:.6, sz:12, c:INK, ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis']);
  tx(s, {x:4.37, y:4.84, w:7.41, h:1.54, sz:18, c:WHITE, a:'center', v:1, f2:WHITE, sd:1}, ['Y']);
  draw(s, [
    ['flowChartConnector', 3.84,4.43,.88,.88, WHITE, {sd:1}],
    [IMG, 4.06,4.66,.43,.43, IC_BLUE],
  ]);
  s.addChart('doughnut', [{ name: 'Sales', labels: ['1st Qtr', '2nd Qtr'], values: [8.2, 3.2] }], {
    x: 4.37, y: 4.84, w: 2.32, h: 1.54, holeSize: 70, chartColors: [BLUE, SNOW],
    showLegend: false, showValue: false, dataBorder: { pt: 0, color: WHITE },
  });
  tx(s, {x:5.23, y:5.46, w:.59, h:.3, sz:12, c:BLUE, f:HEAD, a:'center', nw:1, af:1}, ['60%']);
  draw(s, [
    [LINE, 6.67,4.84,0,1.54, null, {lc:GREY2}],
  ]);
  tx(s, {x:7.03, y:5.14, w:1.68, h:.34, sz:14, c:BLACK, f:HEAD, nw:1, af:1}, ['Your Text Here']);
  tx(s, {x:7.02, y:5.49, w:4.49, h:.6, sz:12, c:INK, ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula dolor aenean massa. ']);
}

// Slide 4 — Four rounded stat bars branching off a vertical pencil
function slide04(s) {
  chrome(s, 4);
  draw(s, [
    [RECT, 5.92,4.44,1.06,.23, BLUE_L, {r:269.6}],
    [RECT, 6.14,4.44,1.06,.23, BLUE, {r:269.6}],
    [RECT, 6.37,4.44,1.06,.23, BLUE_D, {r:269.6}],
    [RECT, 5.91,3.39,1.06,.23, TEAL_L, {r:269.6}],
    [RECT, 6.14,3.38,1.06,.22, TEAL, {r:269.6}],
    [RECT, 6.36,3.38,1.06,.23, TEAL_D, {r:269.6}],
    [P4A, 6.54,1.54,.22,.68, GREY, {r:269.6}],
    [P4B, 5.86,2.31,1.14,.23, PINK_L, {r:269.6}],
    [RECT, 6.12,2.34,1.06,.23, PINK, {r:269.6}],
    [P4C, 6.31,2.3,1.14,.23, PINK_D, {r:269.6}],
    ['ellipse', 6.62,1.75,.05,.19, INK, {r:269.6}],
    [P4D, 6.21,6.12,.97,.68, GREY, {r:269.6}],
    [P4E, 5.89,5.54,1.14,.23, GOLD_L, {r:269.6}],
    [P4E, 6.12,5.54,1.14,.22, GOLD, {r:269.6}],
    [P4E, 6.34,5.54,1.14,.23, GOLD_D, {r:269.6}],
    [P4F, 6.56,6.7,.28,.22, BLACK, {r:269.6}],
    [P4G, 7.23,2.85,3.5,1.02, TEAL, {r:180}],
    [P4H, 7,2.84,.24,1.18, TEAL_XD],
    [P4G, 7.22,4.93,3.5,1.02, GOLD, {r:180}],
    [P4H, 7.02,4.92,.24,1.18, GOLD_XD],
    [P4H, 6.09,1.88,.23,1.27, PINK_XD],
    [P4H, 6.09,4.03,.25,1.21, BLUE_XD],
    [P4G, 2.6,2.01,3.5,1.14, PINK, {fv:1}],
    [P4G, 2.6,4.17,3.5,1.05, BLUE, {fv:1}],
  ]);
  tx(s, {x:2.61, y:.53, w:8.11, h:.77, sz:40, c:INK, f:HEAD, a:'center', af:1}, ['Innovative Pencil ', ['Solutions', {c:BLUE}]]);
  tx(s, {x:3.6, y:2.33, w:2.6, h:.51, sz:12, c:WHITE, af:1}, ['Separated they live in Bookmarksgrove right at.']);
  tx(s, {x:2.77, y:2.29, w:.82, h:.57, sz:28, c:WHITE, f:HEAD, a:'center', nw:1, af:1}, ['17', ['k', {sz:16}]]);
  tx(s, {x:3.6, y:4.45, w:2.6, h:.51, sz:12, c:WHITE, af:1}, ['Separated they live in Bookmarksgrove right at.']);
  tx(s, {x:2.77, y:4.41, w:.82, h:.57, sz:28, c:WHITE, f:HEAD, a:'center', nw:1, af:1}, ['25', ['k', {sz:16}]]);
  tx(s, {x:8.18, y:5.19, w:2.6, h:.51, sz:12, c:WHITE, af:1}, ['Separated they live in Bookmarksgrove right at.']);
  tx(s, {x:7.36, y:5.16, w:.82, h:.57, sz:28, c:WHITE, f:HEAD, a:'center', nw:1, af:1}, ['32', ['k', {sz:16}]]);
  tx(s, {x:8.18, y:3.1, w:2.6, h:.51, sz:12, c:WHITE, af:1}, ['Separated they live in Bookmarksgrove right at.']);
  tx(s, {x:7.36, y:3.07, w:.82, h:.57, sz:28, c:WHITE, f:HEAD, a:'center', nw:1, af:1}, ['19', ['k', {sz:16}]]);
  draw(s, [
    [IMG, 2,2.33,.43,.43, IC_PINK],
    [IMG, 10.96,3.18,.36,.36, IC_TEAL],
    [IMG, 2,4.51,.38,.38, IC_BLUE],
    [IMG, 10.97,5.22,.35,.35, IC_GOLD],
  ]);
}

// Slide 5 — Five pencils of varying height, one per column of copy
function slide05(s) {
  chrome(s, 5);
  draw(s, [
    [P5A, .96,4.41,1.49,1.19, GREY],
    [P5B, 1.7,4.41,.42,1.19, GREY],
    [P5C, 1.29,4.41,.42,1.19, GREY],
    [P5A, 1.49,4.41,.42,.32, CHAR],
    [P5D, 1.59,4.41,.23,.32, CHAR],
    [P5E, 1.59,4.41,.12,.32, CHAR],
    [P5F, 1.95,5.8,.5,1.71, BLUE_D, {r:180}],
    [P5F, 1.46,5.8,.5,1.71, BLUE, {r:180}],
    [P5F, .96,5.8,.5,1.71, BLUE_L, {r:180}],
    [P5G, 1.95,5.43,.5,.77, BLUE_D, {r:180}],
    [P5G, 1.46,5.43,.5,.77, BLUE, {r:180}],
    [P5G, .96,5.43,.5,.77, BLUE_L, {r:180}],
  ]);
  tx(s, {x:.66, y:3.42, w:2.08, h:.81, sz:12, c:BLACK, a:'center', ls:1.2, af:1}, ['Far far', T3]);
  tx(s, {x:.5, y:3.04, w:2.4, h:.37, sz:16, c:BLUE, f:HEAD, a:'center', af:1}, [['Title One ', {b:1}]]);
  draw(s, [
    [P5A, 3.4,3.17,1.5,1.19, GREY],
    [P5B, 4.15,3.17,.42,1.19, GREY],
    [P5H, 3.74,3.17,.41,1.19, GREY],
    [P5A, 3.94,3.17,.42,.32, CHAR],
    [P5A, 4.03,3.17,.23,.32, CHAR],
    [P5I, 4.03,3.17,.12,.32, CHAR],
    [P5J, 4.41,4.17,.5,3.34, PINK_D, {r:180}],
    [P5J, 3.91,4.17,.5,3.34, PINK, {r:180}],
    [P5J, 3.41,4.17,.5,3.34, PINK_L, {r:180}],
    [P5G, 4.41,4.17,.5,.83, PINK_D, {r:180}],
    [P5G, 3.91,4.17,.5,.83, PINK, {r:180}],
    [P5G, 3.41,4.17,.5,.83, PINK_L, {r:180}],
  ]);
  tx(s, {x:3.11, y:2.21, w:2.08, h:.81, sz:12, c:BLACK, a:'center', ls:1.2, af:1}, ['Far far', T3]);
  tx(s, {x:2.94, y:1.84, w:2.4, h:.37, sz:16, c:PINK, f:HEAD, a:'center', af:1}, [['Your Text', {b:1}]]);
  draw(s, [
    [P5A, 5.86,4.62,1.5,1.18, GREY],
    [P5B, 6.59,4.62,.42,1.18, GREY],
    [P5C, 6.18,4.62,.42,1.18, GREY],
    [P5A, 6.38,4.62,.42,.32, CHAR],
    [P5D, 6.48,4.62,.23,.32, CHAR],
    [P5E, 6.48,4.62,.12,.32, CHAR],
    [P5F, 6.86,5.92,.5,1.58, TEAL_D, {r:180}],
    [P5F, 6.36,5.92,.5,1.58, TEAL, {r:180}],
    [P5F, 5.86,5.92,.5,1.58, TEAL_L, {r:180}],
    [P5G, 6.86,5.6,.5,.83, TEAL_D, {r:180}],
    [P5G, 6.36,5.6,.5,.83, TEAL, {r:180}],
    [P5G, 5.86,5.6,.5,.83, TEAL_L, {r:180}],
  ]);
  tx(s, {x:5.55, y:3.53, w:2.08, h:.81, sz:12, c:BLACK, a:'center', ls:1.2, af:1}, ['Far far', T3]);
  tx(s, {x:5.39, y:3.16, w:2.4, h:.37, sz:16, c:TEAL, f:HEAD, a:'center', af:1}, [['Your Text', {b:1}]]);
  draw(s, [
    [P5A, 8.29,4.17,1.49,1.14, GREY],
    [P5B, 9.02,4.17,.42,1.14, GREY],
    [P5H, 8.62,4.17,.41,1.14, GREY],
    [P5A, 8.82,4.17,.42,.32, CHAR],
    [P5A, 8.91,4.17,.23,.32, CHAR],
    [P5I, 8.91,4.17,.12,.32, CHAR],
    [P5F, 9.28,5.42,.5,2.08, GOLD_D, {r:180}],
    [P5F, 8.78,5.42,.5,2.08, GOLD, {r:180}],
    [P5F, 8.29,5.42,.5,2.08, GOLD_L, {r:180}],
    [P5G, 9.28,5.1,.5,.83, GOLD_D, {r:180}],
    [P5G, 8.78,5.1,.5,.83, GOLD, {r:180}],
    [P5G, 8.29,5.1,.5,.83, GOLD_L, {r:180}],
  ]);
  tx(s, {x:8, y:3.07, w:2.08, h:.81, sz:12, c:BLACK, a:'center', ls:1.2, af:1}, ['Far far', T3]);
  tx(s, {x:7.84, y:2.7, w:2.4, h:.37, sz:16, c:GOLD, f:HEAD, a:'center', af:1}, [['Your Text', {b:1}]]);
  draw(s, [
    [P5A, 10.75,2.5,1.49,1.14, GREY],
    [P5B, 11.49,2.5,.42,1.14, GREY],
    [P5H, 11.08,2.5,.41,1.14, GREY],
    [P5A, 11.28,2.5,.42,.32, CHAR],
    [P5A, 11.37,2.5,.23,.32, CHAR],
    [P5I, 11.37,2.5,.12,.32, CHAR],
    [P5J, 11.74,3.77,.5,3.73, SLATE_D, {r:180}],
    [P5J, 11.24,3.77,.5,3.73, SLATE, {r:180}],
    [P5J, 10.75,3.77,.5,3.73, SLATE_L, {r:180}],
    [P5G, 11.74,3.45,.5,.83, SLATE_D, {r:180}],
    [P5G, 11.24,3.45,.5,.83, SLATE, {r:180}],
    [P5G, 10.75,3.45,.5,.83, SLATE_L, {r:180}],
  ]);
  tx(s, {x:10.45, y:1.43, w:2.08, h:.81, sz:12, c:BLACK, a:'center', ls:1.2, af:1}, ['Far far', T3]);
  tx(s, {x:10.28, y:1.06, w:2.4, h:.37, sz:16, c:SLATE, f:HEAD, a:'center', af:1}, [['Your Text', {b:1}]]);
  tx(s, {x:2.61, y:.53, w:8.11, h:.77, sz:40, c:INK, f:HEAD, a:'center', af:1}, ['Creative Pencil ', ['Strategies', {c:BLUE}]]);
}

// Slide 6 — Horizontal pencil timeline with four dashed callouts
function slide06(s) {
  s.background = { color: BLUE };
  chrome(s, 6, true);
  draw(s, [
    ['ellipse', 1.75,3.44,1.31,1.12, GREY],
    [P6A, 11.07,4.04,.49,.57, GREY3, {r:270}],
    [P6B, 11.07,3.4,.49,.57, SNOW, {r:270}],
    [P6C, 11.13,3.72,.38,.57, GREY, {r:270}],
    [P6D, 11.63,3.75,.2,.3, CHAR, {r:270}],
    [P6E, 11.63,3.95,.2,.3, CHAR, {r:270}],
    [P6F, 11.66,3.85,.14,.3, CHAR, {r:270}],
    [P6G, 9.49,2.94,1.12,2.12, PINK_D, {r:270}],
    [P6H, 9.83,2.53,.38,2.2, PINK_L, {r:270}],
    [P6I, 9.82,2.9,.38,2.2, PINK, {r:270}],
    [RECT, 4.98,1.23,.38,4.79, PINK_P, {r:270}],
    [RECT, 4.64,1.95,.38,4.12, PINK, {r:270}],
    [RECT, 2.62,3.22,1.12,1.56, TEAL_D, {r:270}],
    [RECT, 2.42,3.45,.37,.36, INK, {r:270}],
    [RECT, 2.42,3.82,.37,.36, INK_L, {r:270}],
    [RECT, 7.65,3.22,1.12,1.56, TEAL_D, {r:270}],
    [RECT, 5.51,2.64,.38,3.47, PINK_D, {r:270}],
    [RECT, 3.18,2.67,.38,1.92, TEAL_L, {r:270}],
    [RECT, 3.46,2.76,.37,2.48, TEAL, {r:270}],
    [RECT, 8.21,2.67,.38,1.92, TEAL_L, {r:270}],
    [RECT, 8.21,2.49,.37,3.03, TEAL, {r:270}],
    ['flowChartConnector', 3.11,4.95,.59,.59, WHITE, {lc:SNOW,sd:1}],
    [LINE, 3.94,4.2,0,1.45, null, {lc:TEAL,lw:1.5,ld:1}],
  ]);
  tx(s, {x:1.02, y:5.07, w:1.99, h:.61, sz:12, c:WHITE, a:'right', ls:1.3, af:1}, [T1]);
  tx(s, {x:1.9, y:4.77, w:1.11, h:.34, sz:14, c:TEAL, f:HEAD, a:'right', nw:1, af:1}, ['Title One']);
  draw(s, [
    [LINE, 4.35,2.35,0,1.45, null, {lc:PINK,lw:1.5,ld:1}],
    ['flowChartConnector', 4.58,2.61,.59,.59, WHITE, {lc:SNOW,sd:1}],
  ]);
  tx(s, {x:5.29, y:2.74, w:1.99, h:.61, sz:12, c:WHITE, ls:1.3, af:1}, [T1]);
  tx(s, {x:5.29, y:2.44, w:1.07, h:.34, sz:14, c:PINK, f:HEAD, nw:1, af:1}, ['Title Two']);
  draw(s, [
    [LINE, 8.98,4.21,0,1.45, null, {lc:TEAL,lw:1.5,ld:1}],
  ]);
  tx(s, {x:6.09, y:5.07, w:1.99, h:.61, sz:12, c:WHITE, a:'right', ls:1.3, af:1}, [T1]);
  tx(s, {x:6.84, y:4.77, w:1.23, h:.34, sz:14, c:TEAL, f:HEAD, a:'right', nw:1, af:1}, ['Title Three']);
  draw(s, [
    ['flowChartConnector', 8.14,4.94,.59,.59, WHITE, {lc:SNOW,sd:1}],
    [LINE, 9.37,2.35,0,1.45, null, {lc:PINK,lw:1.5,ld:1}],
    ['flowChartConnector', 9.61,2.62,.59,.59, WHITE, {lc:SNOW,sd:1}],
  ]);
  tx(s, {x:10.32, y:2.75, w:1.99, h:.61, sz:12, c:WHITE, ls:1.3, af:1}, [T1]);
  tx(s, {x:10.32, y:2.45, w:1.14, h:.34, sz:14, c:PINK, f:HEAD, nw:1, af:1}, ['Title Four']);
  draw(s, [
    [IMG, 4.72,2.74,.33,.33, IC_PINK],
    [IMG, 8.3,5.1,.27,.27, IC_TEAL],
    [IMG, 9.75,2.76,.3,.3, IC_PINK],
    [IMG, 3.27,5.1,.29,.29, IC_TEAL],
  ]);
  tx(s, {x:2.61, y:.53, w:8.11, h:.77, sz:40, c:WHITE, f:HEAD, a:'center', af:1}, ['Creative Pencil Strategies']);
}

// Slide 7 — Jigsaw pencil built from photo tiles, four labels
function slide07(s) {
  chrome(s, 7);
  draw(s, [
    ['flowChartProcess', 3.51,1.98,.56,.55, PINK, {lc:PINK_W,lw:2}],
    [IMG, 5.71,.82,.96,1.24, 'ACABA6'],
    [P7A, 5.73,4.63,1.85,1.58, GREY],
    [P7B, 6.38,.82,1.24,.96, BLUE, {lc:BLUE_P,lw:2}],
    [P7C, 5.71,1.78,.96,.96, PINK, {lc:PINK_W,lw:2}],
    [P7D, 6.29,5.92,.73,.61, CHAR],
    [P7E, 5.71,3.69,1.25,1.32, TEAL, {lc:TEAL_W,lw:2}],
    [P7F, 6.38,2.73,1.24,1.24, GOLD, {lc:PINK_W,lw:2}],
    [P7G, 6.17,.04,.94,.73, SKIN, {r:330}],
    [P7H, 4.85,-.56,2.87,1.37, SKIN, {r:330}],
    [P7I, 6.81,-.75,1.18,1.05, WHITE, {r:330}],
    [P7J, 6.93,-1.06,2.03,1.36, CHAR, {r:330}],
    ['ellipse', 7.04,-.45,.06,.06, WHITE, {r:330}],
    [P7K, 4.94,.17,.11,.15, SKIN_L, {r:330}],
    [P7L, 5,.44,.14,.15, SKIN_L, {r:330}],
    [P7M, 5.06,.77,.12,.14, SKIN_L, {r:330}],
    [P7N, 5.53,1.13,.18,.18, SKIN_L, {r:285}],
    [IMG, 5.71,2.45,.96,1.52, '9A93A5'],
    [IMG, 6.38,1.5,1.24,1.52, '90978B'],
    [IMG, 6.67,3.69,.93,1.21, '7F7B69'],
    ['flowChartProcess', 9.26,1.02,.56,.55, BLUE, {lc:BLUE_P,lw:2}],
  ]);
  tx(s, {x:9.89, y:1.27, w:2.54, h:.61, sz:12, c:BLACK, ls:1.3, af:1}, [T2]);
  tx(s, {x:9.89, y:.96, w:1.26, h:.34, sz:14, c:BLUE, f:HEAD, nw:1, af:1}, ['Title Three']);
  draw(s, [
    [IMG, 3.64,2.1,.31,.31, IC_WHITE],
    ['flowChartProcess', 3.51,4.01,.56,.55, TEAL, {lc:TEAL_W,lw:2}],
    [IMG, 3.65,4.15,.28,.28, IC_WHITE],
    [IMG, 9.41,1.17,.25,.25, IC_WHITE],
    ['flowChartProcess', 9.26,2.89,.56,.55, GOLD, {lc:GOLD_W,lw:2}],
  ]);
  tx(s, {x:9.89, y:3.14, w:2.54, h:.61, sz:12, c:BLACK, ls:1.3, af:1}, [T2]);
  tx(s, {x:9.89, y:2.84, w:1.14, h:.34, sz:14, c:GOLD, f:HEAD, nw:1, af:1}, ['Title Four']);
  draw(s, [
    [IMG, 9.42,3.04,.25,.25, IC_WHITE],
  ]);
  tx(s, {x:.92, y:4.21, w:2.54, h:.61, sz:12, c:BLACK, a:'right', ls:1.3, af:1}, [T2]);
  tx(s, {x:2.33, y:3.91, w:1.12, h:.34, sz:14, c:TEAL, f:HEAD, a:'right', nw:1, af:1}, ['Title Two']);
  tx(s, {x:.92, y:2.18, w:2.54, h:.61, sz:12, c:BLACK, a:'right', ls:1.3, af:1}, [T2]);
  tx(s, {x:2.29, y:1.87, w:1.15, h:.34, sz:14, c:PINK, f:HEAD, a:'right', nw:1, af:1}, ['Title One ']);
  draw(s, [
    [LINE, 4.15,4.29,1.56,.02, null, {r:180,lc:GREY,lw:1.5,ld:1}],
    [LINE, 4.14,2.25,1.57,0, null, {r:180,lc:GREY,lw:1.5,ld:1}],
    [LINE, 7.62,1.3,1.57,0, null, {r:180,lc:GREY,lw:1.5,ld:1}],
    [LINE, 7.62,3.17,1.57,0, null, {r:180,lc:GREY,lw:1.5,ld:1}],
  ]);
}

// Slide 8 — Headline with a floating stat card and a bullet card
function slide08(s) {
  chrome(s, 8);
  tx(s, {x:.89, y:4.09, w:5.54, h:1.92, sz:36, c:BLACK, f:HEAD, af:1}, ['Draw Your ', ['Future Creative Pencil ', {c:BLUE}], 'Strategies That Sell']);
  draw(s, [
    [RECT, 5.83,2.18,1.99,2.38, WHITE, {sd:1}],
  ]);
  tx(s, {x:6, y:2.46, w:1.15, h:.34, sz:14, c:BLUE, f:HEAD, nw:1, af:1}, ['Your Text']);
  tx(s, {x:6.01, y:2.8, w:1.65, h:.8, sz:11, c:INK, ls:1.3, af:1}, ['Lorem ipsum dolor sit met consectetuer adipiscing elit. ']);
  tx(s, {x:6.03, y:3.83, w:.94, h:.44, sz:20, c:BLACK, f:HEAD, af:1}, ['+100']);
  tx(s, {x:6.8, y:3.95, w:.82, h:.27, sz:10, c:GREY4, nw:1, af:1}, ['Followers']);
  draw(s, [
    [RECT, 8.57,4.26,3.41,1.45, WHITE, {sd:1}],
    [IMG, 8.83,4.54,.23,.23, IC_BLUE],
  ]);
  tx(s, {x:9.04, y:4.51, w:2.25, h:.29, sz:11, c:GREY4, af:1}, ['Lorem ipsum dolor sit amet. ']);
  draw(s, [
    [IMG, 8.83,4.87,.23,.23, IC_BLUE],
  ]);
  tx(s, {x:9.04, y:4.84, w:2.75, h:.29, sz:11, c:GREY4, af:1}, ['Aenean commodo ligula eget dolor.']);
  draw(s, [
    [IMG, 8.83,5.21,.23,.23, IC_BLUE],
  ]);
  tx(s, {x:9.04, y:5.18, w:2.25, h:.29, sz:11, c:GREY4, af:1}, ['Behind the word mountains.']);
}

// Slide 9 — Three-icon feature card above a photo strip
function slide09(s) {
  chrome(s, 9);
  tx(s, {x:1.45, y:.86, w:10.44, h:.77, sz:40, c:BLUE, f:HEAD, a:'center', af:1}, [['Our Best ', {c:BLACK}], 'Category Creative ', ['Pencil', {c:BLACK}]]);
  tx(s, {x:2.61, y:1.88, w:7.89, h:.6, sz:12, c:INK, a:'center', ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Cum sociis natoque penatibus et magnis dis parturient montes.']);
  draw(s, [
    [RECT, 1.8,2.99,9.71,2.18, WHITE, {sd:1}],
  ]);
  tx(s, {x:2.8, y:3.94, w:1.15, h:.34, sz:14, c:BLACK, f:HEAD, a:'center', nw:1, af:1}, ['Your Text']);
  tx(s, {x:2.19, y:4.28, w:2.39, h:.56, sz:11, c:INK, a:'center', ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit.']);
  draw(s, [
    ['flowChartConnector', 3.12,3.32,.55,.55, BLUE],
    [IMG, 3.26,3.46,.26,.26, IC_WHITE],
  ]);
  tx(s, {x:6.08, y:3.94, w:1.15, h:.34, sz:14, c:BLACK, f:HEAD, a:'center', nw:1, af:1}, ['Your Text']);
  tx(s, {x:5.47, y:4.28, w:2.39, h:.56, sz:11, c:INK, a:'center', ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ']);
  tx(s, {x:9.35, y:3.94, w:1.15, h:.34, sz:14, c:BLACK, f:HEAD, a:'center', nw:1, af:1}, ['Your Text']);
  tx(s, {x:8.73, y:4.28, w:2.39, h:.56, sz:11, c:INK, a:'center', ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. ']);
  draw(s, [
    ['flowChartConnector', 6.37,3.32,.55,.55, PINK],
    ['flowChartConnector', 9.63,3.32,.55,.55, TEAL],
    [IMG, 6.51,3.45,.3,.3, IC_WHITE],
    [IMG, 9.79,3.45,.28,.28, IC_WHITE],
  ]);
}

// Slide 10 — S-curve journey with four upright pencils
function slide10(s) {
  chrome(s, 10);
  draw(s, [
    ['arc', 1.81,3.1,2.43,2.43, null, {r:135,lc:GREY2,lw:4.5,a1:223.86,a2:6.85}],
    ['arc', 4.24,3.08,2.43,2.43, null, {r:180,fv:1,lc:GREY2,lw:4.5,a1:178.44,a2:0}],
    ['arc', 6.67,3.1,2.43,2.43, null, {r:45,fv:1,lc:GREY2,lw:4.5,a1:223.86,a2:50.56}],
    ['arc', 9.09,2.91,2.43,2.43, null, {r:225,fv:1,lc:GREY2,lw:4.5,a1:264.24,a2:48.33}],
    [P4D, 1.9,4.68,.5,.35, GREY, {r:90.4,fv:1}],
    [P4E, 1.73,5.21,.59,.12, BLUE_L, {r:90.4,fv:1}],
    [P4E, 1.85,5.21,.59,.12, BLUE, {r:90.4,fv:1}],
    [P4E, 1.97,5.21,.59,.12, BLUE_D, {r:90.4,fv:1}],
    [P4F, 2.08,4.62,.14,.11, BLACK, {r:90.4,fv:1}],
    [P4A, 2.08,5.78,.11,.35, GREY, {r:90.4,fv:1}],
    [P4B, 1.73,5.62,.59,.12, BLUE_L, {r:90.4,fv:1}],
    [RECT, 1.87,5.6,.55,.12, BLUE, {r:90.4,fv:1}],
    [P4C, 1.96,5.62,.59,.12, BLUE_D, {r:90.4,fv:1}],
    ['ellipse', 2.12,5.92,.03,.1, INK, {r:90.4,fv:1}],
    [P4D, 5.21,3.4,.5,.35, GREY, {r:269.6}],
    [P4E, 5.05,3.1,.59,.12, PINK_L, {r:269.6}],
    [P4E, 5.16,3.1,.59,.12, PINK, {r:269.6}],
    [P4E, 5.28,3.09,.59,.12, PINK_D, {r:269.6}],
    [P4F, 5.39,3.69,.14,.11, BLACK, {r:269.6}],
    [P4A, 5.39,2.29,.11,.35, GREY, {r:269.6}],
    [P4B, 5.04,2.68,.59,.12, PINK_L, {r:269.6}],
    [RECT, 5.18,2.7,.55,.12, PINK, {r:269.6}],
    [P4C, 5.28,2.68,.59,.12, PINK_D, {r:269.6}],
    ['ellipse', 5.44,2.4,.03,.1, INK, {r:269.6}],
    [P4D, 7.66,4.73,.5,.35, GREY, {r:90.4,fv:1}],
    [P4E, 7.5,5.26,.59,.12, TEAL_L, {r:90.4,fv:1}],
    [P4E, 7.62,5.26,.59,.12, TEAL, {r:90.4,fv:1}],
    [P4E, 7.73,5.26,.59,.12, TEAL_D, {r:90.4,fv:1}],
    [P4F, 7.84,4.67,.14,.11, BLACK, {r:90.4,fv:1}],
    [P4A, 7.85,5.83,.11,.35, GREY, {r:90.4,fv:1}],
    [P4B, 7.5,5.67,.59,.12, TEAL_L, {r:90.4,fv:1}],
    [RECT, 7.63,5.66,.55,.12, TEAL, {r:90.4,fv:1}],
    [P4C, 7.73,5.68,.59,.12, TEAL_D, {r:90.4,fv:1}],
    ['ellipse', 7.89,5.98,.03,.1, INK, {r:90.4,fv:1}],
    [P4D, 10.89,3.4,.5,.35, GREY, {r:269.6}],
    [P4E, 10.73,3.1,.59,.12, GOLD_L, {r:269.6}],
    [P4E, 10.85,3.1,.59,.12, GOLD, {r:269.6}],
    [P4E, 10.96,3.09,.59,.12, GOLD_D, {r:269.6}],
    [P4F, 11.07,3.69,.14,.11, BLACK, {r:269.6}],
    [P4A, 11.08,2.29,.11,.35, GREY, {r:269.6}],
    [P4B, 10.73,2.68,.59,.12, GOLD_L, {r:269.6}],
    [RECT, 10.87,2.7,.55,.12, GOLD, {r:269.6}],
    [P4C, 10.96,2.68,.59,.12, GOLD_D, {r:269.6}],
    ['ellipse', 11.12,2.4,.03,.1, INK, {r:269.6}],
    ['flowChartConnector', 1.88,5.66,.51,.51, WHITE, {fv:1,sd:1}],
    ['flowChartConnector', 7.65,5.66,.51,.51, WHITE, {fv:1,sd:1}],
    ['flowChartConnector', 5.19,2.15,.51,.51, WHITE, {fv:1,sd:1}],
    ['flowChartConnector', 10.89,2.15,.51,.51, WHITE, {fv:1,sd:1}],
    [IMG, 5.33,2.28,.27,.27, IC_PINK],
    [IMG, 2.01,5.8,.26,.26, IC_BLUE],
    [IMG, 7.79,5.8,.26,.26, IC_TEAL],
    [IMG, 11.04,2.29,.22,.22, IC_GOLD],
  ]);
  tx(s, {x:1.12, y:3.17, w:2.01, h:.87, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, [T2]);
  tx(s, {x:1.59, y:2.87, w:1.07, h:.34, sz:14, c:BLUE, f:HEAD, a:'center', nw:1, af:1}, ['Title One']);
  tx(s, {x:4.47, y:4.54, w:2.01, h:.87, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, [T2]);
  tx(s, {x:4.94, y:4.24, w:1.07, h:.34, sz:14, c:PINK, f:HEAD, a:'center', nw:1, af:1}, ['Ttile Two']);
  tx(s, {x:6.91, y:3.21, w:2.01, h:.87, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, [T2]);
  tx(s, {x:7.29, y:2.9, w:1.23, h:.34, sz:14, c:TEAL, f:HEAD, a:'center', nw:1, af:1}, ['Title Three']);
  tx(s, {x:10.12, y:4.54, w:2.01, h:.87, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, [T2]);
  tx(s, {x:10.57, y:4.24, w:1.1, h:.34, sz:14, c:GOLD, f:HEAD, a:'center', nw:1, af:1}, ['Ttile Four']);
  tx(s, {x:2.17, y:.54, w:9, h:.77, sz:44, c:BLUE, f:HEAD, a:'center', ls:.9, sb:10}, ['Creative Pencil ', ['Infographic', {c:BLACK}]]);
}

// Slide 11 — Four horizontal pencils alternating left and right
function slide11(s) {
  s.background = { color: BLUE };
  chrome(s, 11, true);
  draw(s, [
    [P11A, 5.1,4.27,.7,1.15, GREY, {r:90}],
    [P11A, 5.1,4.27,.7,1.15, null, {r:90}],
    [P11B, 2.36,2.31,.29,5.05, TEAL, {r:90}],
    [P11B, 2.4,2.57,.24,5.06, GOLD_O, {r:90}],
    [P11B, 2.4,2.05,.23,5.06, TEAL_L, {r:90}],
    [P11C, 5.67,4.63,.3,.44, CHAR, {r:90}],
    [P11D, 5.1,4.66,.37,.7, GREY, {r:90}],
    [P11E, 2.4,2.57,.24,5.06, TEAL_D, {r:90}],
    [P11F, 5.75,4.7,.17,.42, CHAR, {r:90}],
    [P11A, 7.51,2.94,.7,1.15, GREY, {r:90,fv:1}],
    [P11A, 7.51,2.94,.7,1.15, null, {r:90,fv:1}],
    [P11B, 10.66,.98,.29,5.05, PINK, {r:90,fv:1}],
    [P11B, 10.68,1.24,.24,5.06, GOLD_O, {r:90,fv:1}],
    [P11B, 10.69,.72,.23,5.06, PINK_L, {r:90,fv:1}],
    [P11C, 7.34,3.29,.3,.44, CHAR, {r:90,fv:1}],
    [P11D, 7.85,3.33,.37,.7, GREY, {r:90,fv:1}],
    [P11G, 10.73,1.05,.16,5.05, PINK, {r:90,fv:1}],
    [P11E, 10.68,1.24,.24,5.06, PINK_D, {r:90,fv:1}],
    [P11F, 7.4,3.37,.17,.42, CHAR, {r:90,fv:1}],
    [P11A, 5.12,1.57,.7,1.15, GREY, {r:90}],
    [P11A, 5.12,1.57,.7,1.15, null, {r:90}],
    [P11B, 2.38,-.39,.29,5.05, TEAL, {r:90}],
    [P11B, 2.41,-.13,.24,5.06, GOLD_O, {r:90}],
    [P11B, 2.42,-.65,.23,5.06, TEAL_L, {r:90}],
    [P11C, 5.69,1.93,.3,.44, CHAR, {r:90}],
    [P11D, 5.12,1.96,.37,.7, GREY, {r:90}],
    [P11E, 2.41,-.13,.24,5.06, TEAL_D, {r:90}],
    [P11F, 5.76,2,.17,.42, CHAR, {r:90}],
    [P11A, 7.52,5.63,.7,1.15, GREY, {r:90,fv:1}],
    [P11A, 7.52,5.63,.7,1.15, null, {r:90,fv:1}],
    [P11B, 10.67,3.68,.29,5.05, PINK, {r:90,fv:1}],
    [P11B, 10.69,3.94,.24,5.06, GOLD_O, {r:90,fv:1}],
    [P11B, 10.7,3.42,.23,5.06, PINK_L, {r:90,fv:1}],
    [P11C, 7.36,5.99,.3,.44, CHAR, {r:90,fv:1}],
    [P11D, 7.86,6.03,.37,.7, GREY, {r:90,fv:1}],
    [P11E, 10.69,3.94,.24,5.06, PINK_D, {r:90,fv:1}],
    [P11F, 7.42,6.07,.17,.42, CHAR, {r:90,fv:1}],
    ['flowChartConnector', 6.97,1.83,.51,.51, TEAL, {fv:1,sd:1}],
    [IMG, 7.1,1.97,.25,.25, IC_WHITE],
  ]);
  tx(s, {x:7.58, y:1.97, w:2.54, h:.61, sz:12, c:WHITE, ls:1.3, af:1}, [T2]);
  tx(s, {x:7.58, y:1.67, w:1.11, h:.34, sz:14, c:TEAL, f:HEAD, nw:1, af:1}, ['Title One']);
  tx(s, {x:3.22, y:3.33, w:2.54, h:.61, sz:12, c:WHITE, a:'right', ls:1.3, af:1}, [T2]);
  tx(s, {x:4.68, y:3.02, w:1.07, h:.34, sz:14, c:PINK, f:HEAD, a:'right', nw:1, af:1}, ['Title Two']);
  draw(s, [
    ['flowChartConnector', 5.85,3.24,.51,.51, PINK, {fv:1,sd:1}],
    [IMG, 5.99,3.36,.26,.26, IC_WHITE],
    ['flowChartConnector', 6.97,4.6,.51,.51, TEAL, {fv:1,sd:1}],
    [IMG, 7.12,4.73,.24,.24, IC_WHITE],
  ]);
  tx(s, {x:7.58, y:4.7, w:2.54, h:.61, sz:12, c:WHITE, ls:1.3, af:1}, [T2]);
  tx(s, {x:7.58, y:4.4, w:1.23, h:.34, sz:14, c:TEAL, f:HEAD, nw:1, af:1}, ['Title Three']);
  draw(s, [
    ['flowChartConnector', 5.85,5.94,.51,.51, PINK, {fv:1,sd:1}],
    [IMG, 5.98,6.08,.24,.24, IC_WHITE],
  ]);
  tx(s, {x:3.23, y:6.05, w:2.54, h:.61, sz:12, c:WHITE, a:'right', ls:1.3, af:1}, [T2]);
  tx(s, {x:4.62, y:5.75, w:1.14, h:.34, sz:14, c:PINK, f:HEAD, a:'right', nw:1, af:1}, ['Title Four']);
  tx(s, {x:2.17, y:.54, w:9, h:.77, sz:40, c:WHITE, f:HEAD, a:'center', ls:.9, sb:10}, ['Pencil Driven Infographic']);
}

// Slide 12 — Pencil "flower" of five petals with outer labels
function slide12(s) {
  chrome(s, 12);
  draw(s, [
    [P12A, 4.06,4.39,2.56,1.27, BLUE],
    [P12B, 4.64,3.01,1.97,2.32, PINK],
    [P12C, 5.97,2.21,1.27,2.81, TEAL],
    [P12D, 6.69,4.39,2.56,1.27, SLATE],
    [P12E, 6.68,3.08,1.97,2.32, GOLD],
    [P12F, 6.11,5.72,1.1,1.64, GREY, {r:180,fv:1}],
    [P12G, 6.42,5.65,.44,.52, CHAR, {r:180,fv:1}],
    [P12H, 6.42,6.11,.44,.06, SKIN_M, {r:180,fv:1}],
    [P12I, 6.87,6.73,.35,.79, SLATE, {r:180,fv:1}],
    [P12J, 6.1,6.74,.38,.78, BLUE, {r:180,fv:1}],
    [P12K, 6.46,6.81,.43,.71, TEAL, {r:180,fv:1}],
  ]);
  tx(s, {x:9.91, y:4.87, w:2.89, h:.63, sz:12, c:BLACK, ls:1.2, af:1}, ['Separated they live in Bookmarksgrove right at the.']);
  tx(s, {x:9.91, y:4.52, w:2.56, h:.34, sz:14, c:SLATE, f:HEAD, af:1}, ['Title Five']);
  draw(s, [
    ['ellipse', 9.61,4.67,.18,.18, SLATE],
  ]);
  tx(s, {x:8.91, y:2.96, w:2.89, h:.63, sz:12, c:BLACK, ls:1.2, af:1}, ['Separated they live in Bookmarksgrove right at the.']);
  tx(s, {x:8.91, y:2.61, w:2.56, h:.34, sz:14, c:GOLD, f:HEAD, af:1}, ['Title Four']);
  draw(s, [
    ['ellipse', 8.61,2.75,.18,.18, GOLD],
  ]);
  tx(s, {x:5.22, y:1.42, w:2.89, h:.57, sz:12, c:BLACK, a:'center', ls:1.2, af:1}, ['Separated they live in Bookmarksgrove right at the.']);
  tx(s, {x:5.38, y:1.06, w:2.56, h:.34, sz:14, c:TEAL, f:HEAD, a:'center', af:1}, ['Title Three']);
  draw(s, [
    ['ellipse', 6.58,.79,.18,.18, TEAL],
  ]);
  tx(s, {x:.48, y:4.88, w:2.89, h:.57, sz:12, c:BLACK, a:'right', ls:1.2, af:1}, ['Separated they live in Bookmarksgrove right at the.']);
  tx(s, {x:.81, y:4.53, w:2.56, h:.34, sz:14, c:BLUE, f:HEAD, a:'right', af:1}, ['Title One']);
  draw(s, [
    ['ellipse', 3.48,4.68,.18,.18, BLUE],
  ]);
  tx(s, {x:1.2, y:2.96, w:2.89, h:.63, sz:12, c:BLACK, a:'right', ls:1.2, af:1}, ['Separated they live in Bookmarksgrove right at the.']);
  tx(s, {x:1.53, y:2.61, w:2.56, h:.34, sz:14, c:PINK, f:HEAD, a:'right', af:1}, ['Title Two']);
  draw(s, [
    ['ellipse', 4.21,2.75,.18,.18, PINK],
    [IMG, 5.04,3.39,.47,.47, IC_PINK],
    [IMG, 6.43,2.61,.42,.42, IC_TEAL],
    [IMG, 4.41,4.8,.43,.43, IC_BLUE],
    [IMG, 7.83,3.49,.39,.39, IC_GOLD],
    [IMG, 8.41,4.8,.43,.43, IC_SLATE],
  ]);
}

// Slide 13 — Vertical pencil with four looped connectors
function slide13(s) {
  s.background = { color: BLUE };
  chrome(s, 13, true);
  draw(s, [
    ['ellipse', 6.41,1.59,.55,.51, SNOW, {r:90}],
    [P6A, 6.72,6.4,.22,.16, SNOW, {r:180,fv:1}],
    [P6B, 6.43,6.4,.22,.16, GREY3, {r:180,fv:1}],
    [P6C, 6.6,6.4,.17,.16, GREY, {r:180,fv:1}],
    [P6D, 6.59,6.55,.09,.09, CHAR2, {r:180,fv:1}],
    [P6E, 6.68,6.55,.09,.09, CHAR4, {r:180,fv:1}],
    [P6F, 6.65,6.55,.07,.09, CHAR3, {r:180,fv:1}],
    [P13A, 6.43,5.91,.51,.51, SLATE_L, {r:180,fv:1}],
    [P6H, 6.43,5.91,.17,.51, SLATE_D, {r:180,fv:1}],
    [P6I, 6.6,5.91,.17,.51, SLATE, {r:180,fv:1}],
    [RECT, 6.43,1.82,.51,4.09, SLATE_L, {r:180,fv:1}],
    [RECT, 6.43,1.82,.17,4.09, SLATE_D, {r:180,fv:1}],
    [RECT, 6.6,1.82,.17,4.09, SLATE, {r:180,fv:1}],
    [RECT, 6.43,2.1,.63,.2, PINK_P],
    [RECT, 6.64,2.1,.21,.2, PINK_L],
    [RECT, 6.84,2.1,1.23,.2, PINK],
    ['blockArc', 5.99,2.1,.89,.89, PINK_W, {r:90,fv:1,a1:179.99,a2:0,a3:.23}],
    ['blockArc', 6.5,2.79,.89,.89, TEAL_W, {r:90,a1:179.99,a2:0,a3:.23}],
    [RECT, 6.31,3.47,.63,.2, TEAL_P],
    [RECT, 6.52,3.47,.21,.2, TEAL_L],
    [RECT, 5.29,3.47,1.23,.2, TEAL],
    [RECT, 6.43,4.56,.63,.2, PINK_P],
    [RECT, 6.64,4.56,.21,.2, PINK_L],
    [RECT, 6.84,4.55,1.23,.2, PINK],
    ['blockArc', 5.99,4.55,.89,.89, PINK_W, {r:90,fv:1,a1:179.99,a2:0,a3:.23}],
    ['blockArc', 6.5,5.24,.89,.89, TEAL_W, {r:90,a1:179.99,a2:0,a3:.23}],
    [RECT, 6.31,5.92,.63,.2, TEAL_P],
    [RECT, 6.52,5.92,.21,.2, TEAL_L],
    [RECT, 5.29,5.92,1.23,.2, TEAL],
    ['flowChartConnector', 7.92,1.93,.55,.55, PINK, {r:90,lc:SNOW,sd:1}],
    ['flowChartConnector', 7.92,4.38,.55,.55, PINK, {r:90,lc:SNOW,sd:1}],
    ['flowChartConnector', 4.87,3.29,.55,.55, TEAL, {r:90,lc:SNOW,sd:1}],
    ['flowChartConnector', 4.87,5.75,.55,.55, TEAL, {r:90,lc:SNOW,sd:1}],
  ]);
  tx(s, {x:8.65, y:2.07, w:2.07, h:.61, sz:12, c:WHITE, ls:1.3, af:1}, [T1]);
  tx(s, {x:8.65, y:1.79, w:1.04, h:.33, sz:14, c:PINK, f:HEAD, nw:1, af:1}, ['Title One']);
  tx(s, {x:8.65, y:4.58, w:2.08, h:.61, sz:12, c:WHITE, ls:1.3, af:1}, [T1]);
  tx(s, {x:8.65, y:4.3, w:1.26, h:.34, sz:14, c:PINK, f:HEAD, nw:1, af:1}, ['Title Three']);
  tx(s, {x:2.62, y:5.95, w:2.09, h:.61, sz:12, c:WHITE, a:'right', ls:1.3, af:1}, [T1]);
  tx(s, {x:3.56, y:5.67, w:1.14, h:.34, sz:14, c:TEAL, f:HEAD, a:'right', nw:1, af:1}, ['Title Four']);
  tx(s, {x:2.61, y:3.5, w:2.08, h:.61, sz:12, c:WHITE, a:'right', ls:1.3, af:1}, [T1]);
  tx(s, {x:3.65, y:3.22, w:1.04, h:.33, sz:14, c:TEAL, f:HEAD, a:'right', nw:1, af:1}, ['Title Two']);
  draw(s, [
    [IMG, 8.08,2.08,.27,.27, IC_WHITE],
    [IMG, 5.03,5.91,.22,.22, IC_WHITE],
    [IMG, 5.03,3.45,.26,.26, IC_WHITE],
    [IMG, 8.08,4.54,.23,.23, IC_WHITE],
  ]);
  tx(s, {x:2.17, y:.54, w:9, h:.77, sz:40, c:WHITE, f:HEAD, ls:.9, sb:10}, ['The Power of Pencil Solutions']);
}

// Slide 14 — Horizontal pencil crossed by five chevrons
function slide14(s) {
  chrome(s, 14);
  draw(s, [
    [P6A, 10.61,3.81,.32,.3, SNOW, {r:90,fv:1}],
    [P6B, 10.61,4.23,.32,.3, GREY3, {r:90,fv:1}],
    [P6C, 10.65,4.02,.25,.3, GREY, {r:90,fv:1}],
    [P6D, 10.92,4.15,.13,.16, CHAR2, {r:90,fv:1}],
    [P6E, 10.92,4.02,.13,.16, CHAR4, {r:90,fv:1}],
    [P6F, 10.94,4.09,.1,.16, CHAR3, {r:90,fv:1}],
    [P14A, 9.83,3.7,.74,.93, BLUE_L, {r:90,fv:1}],
    [P6H, 10.08,3.95,.25,.93, BLUE_D, {r:90,fv:1}],
    [P6I, 10.07,3.7,.25,.93, BLUE, {r:90,fv:1}],
    [RECT, 5.63,.43,.74,7.46, BLUE_L, {r:90,fv:1}],
    [RECT, 5.88,.68,.25,7.46, BLUE_D, {r:90,fv:1}],
    [RECT, 5.88,.43,.25,7.46, BLUE, {r:90,fv:1}],
    [P5A, 3.02,3.58,.21,.21, PINK_D],
    [P14B, 4.49,4.53,.21,.21, PINK_D],
    [P14C, 3.12,3.58,1.49,1.16, PINK],
    [P14C, 4.72,2.65,1.49,1.16, TEAL_D],
    [P5A, 5.5,2.65,.21,.21, TEAL_D, {r:180,fv:1}],
    [P14C, 4.12,2.65,1.49,1.16, TEAL, {r:180,fv:1}],
    [P14B, 6.97,5.48,.21,.21, GOLD_D],
    [P14C, 5.6,4.53,1.49,1.16, GOLD],
    [P14C, 7.07,2.65,1.49,1.16, PURPLE, {r:180,fv:1}],
    [P14C, 6.17,4.53,1.49,1.16, GOLD_D, {r:180,fv:1}],
    [P5A, 9.99,3.58,.21,.21, SLATE_D, {r:180,fv:1}],
    [P14B, 8.51,4.53,.21,.21, SLATE_D, {r:180,fv:1}],
    [P14C, 8.61,3.58,1.49,1.16, SLATE, {r:180,fv:1}],
    [P14C, 7.64,2.64,1.49,1.16, PURPLE_D],
    [IMG, 3.69,3.94,.41,.41, IC_WHITE],
    [IMG, 4.5,1.23,.41,.41, IC_WHITE],
    [IMG, 7.51,3.16,.35,.35, IC_WHITE],
    [IMG, 4.61,3.09,.37,.37, IC_WHITE],
    [IMG, 6.04,4.67,.36,.36, IC_WHITE],
    [IMG, 9.15,3.94,.41,.41, IC_WHITE],
  ]);
  tx(s, {x:2.71, y:5.17, w:1.99, h:.61, sz:12, c:BLACK, a:'right', ls:1.3, af:1}, [T1]);
  tx(s, {x:3.67, y:4.88, w:1.02, h:.32, sz:14, c:PINK, f:HEAD, a:'right', nw:1, af:1}, ['Title One']);
  tx(s, {x:3.67, y:1.94, w:2.03, h:.61, sz:12, c:BLACK, a:'right', ls:1.3, af:1}, [T1]);
  tx(s, {x:4.68, y:1.65, w:1.02, h:.32, sz:14, c:TEAL, f:HEAD, a:'right', nw:1, af:1}, ['Title Two']);
  tx(s, {x:7.59, y:1.94, w:2.03, h:.61, sz:12, c:BLACK, ls:1.3, af:1}, [T1]);
  tx(s, {x:7.59, y:1.65, w:1.05, h:.32, sz:14, c:PURPLE, f:HEAD, nw:1, af:1}, ['Title Four']);
  tx(s, {x:8.56, y:5.17, w:2.06, h:.61, sz:12, c:BLACK, ls:1.3, af:1}, [T1]);
  tx(s, {x:8.56, y:4.88, w:1.01, h:.32, sz:14, c:SLATE, f:HEAD, nw:1, af:1}, ['Title Five']);
  tx(s, {x:5.6, y:6.12, w:2.15, h:.61, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, [T1]);
  tx(s, {x:6.09, y:5.84, w:1.17, h:.32, sz:14, c:GOLD, f:HEAD, a:'center', nw:1, af:1}, ['Title Three']);
  tx(s, {x:2.17, y:.54, w:9, h:.77, sz:40, c:BLACK, f:HEAD, a:'center', ls:.9, sb:10}, [['The Power ', {c:BLUE}], 'of', [' ', {c:BLUE}], 'Pencil Solutions']);
}

// Slide 15 — Pixel-block pencil arrow with six captions
function slide15(s) {
  chrome(s, 15);
  draw(s, [
    [P6A, 9.15,3.69,.66,.52, GREY, {r:270}],
    [P6B, 9.15,2.84,.66,.52, GREY, {r:270}],
    [P6C, 9.23,3.26,.5,.52, GREY, {r:270}],
    [P6D, 9.73,3.25,.27,.28, CHAR, {r:270}],
    [P6E, 9.73,3.52,.27,.27, CHAR, {r:270}],
    [P6F, 9.77,3.39,.19,.28, CHAR, {r:270}],
    [RECT, 5.07,3.53,1.51,.99, TEAL_L, {r:270}],
    [RECT, 5.58,3.03,.5,.99, TEAL_D, {r:270}],
    [RECT, 5.58,3.54,.5,.99, TEAL, {r:270}],
    [P6G, 8.03,3.02,1.51,.99, SLATE_L, {r:270}],
    [P6H, 8.54,2.52,.5,.99, SLATE_D, {r:270}],
    [P6I, 8.54,3.02,.5,.99, SLATE, {r:270}],
    [RECT, 4.09,3.03,1.51,.99, PINK_L, {r:270}],
    [RECT, 4.58,2.53,.5,.99, PINK_D, {r:270}],
    [RECT, 4.58,3.03,.5,.99, PINK, {r:270}],
    [RECT, 3.1,3.52,1.51,.99, BLUE_L, {r:270}],
    [RECT, 3.6,3.02,.5,.99, BLUE_D, {r:270}],
    [RECT, 3.6,3.52,.5,.99, BLUE, {r:270}],
    [RECT, 6.05,3.02,1.51,.99, GOLD_L, {r:270}],
    [RECT, 6.55,2.52,.5,.99, GOLD_D, {r:270}],
    [RECT, 6.55,3.02,.5,.99, GOLD, {r:270}],
    [RECT, 7.04,3.52,1.51,.99, PURPLE_L, {r:270}],
    [RECT, 7.54,3.02,.5,.99, PURPLE_D, {r:270}],
    [RECT, 7.54,3.52,.5,.99, PURPLE, {r:270}],
    ['flowChartConnector', 1.83,5.26,.6,.6, WHITE, {fv:1,sd:1}],
  ]);
  tx(s, {x:1.23, y:5.85, w:1.8, h:.61, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, ['Separated they live in Bookmarksgrove']);
  draw(s, [
    [IMG, 2,5.44,.26,.26, IC_BLUE],
    ['flowChartConnector', 5.72,5.27,.6,.6, WHITE, {fv:1,sd:1}],
    [IMG, 5.91,5.44,.26,.26, IC_TEAL],
  ]);
  tx(s, {x:5.09, y:5.85, w:1.8, h:.61, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, ['Separated they live in Bookmarksgrove']);
  draw(s, [
    ['flowChartConnector', 9.54,5.25,.6,.6, WHITE, {fv:1,sd:1}],
  ]);
  tx(s, {x:8.95, y:5.85, w:1.8, h:.61, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, ['Separated they live in Bookmarksgrove']);
  draw(s, [
    [IMG, 9.72,5.43,.25,.25, IC_PURPLE],
    ['flowChartConnector', 3.13,1.02,.59,.59, WHITE, {fv:1,sd:1}],
  ]);
  tx(s, {x:2.5, y:1.59, w:1.8, h:.61, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, ['Separated they live in Bookmarksgrove']);
  draw(s, [
    [IMG, 3.33,1.2,.27,.27, IC_PINK],
    ['flowChartConnector', 6.97,1,.59,.59, WHITE, {fv:1,sd:1}],
  ]);
  tx(s, {x:6.36, y:1.59, w:1.8, h:.61, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, ['Separated they live in Bookmarksgrove']);
  draw(s, [
    [IMG, 7.15,1.18,.23,.23, IC_GOLD],
    ['flowChartConnector', 10.83,1.03,.59,.59, WHITE, {fv:1,sd:1}],
  ]);
  tx(s, {x:10.22, y:1.62, w:1.8, h:.61, sz:12, c:BLACK, a:'center', ls:1.3, af:1}, ['Separated they live in Bookmarksgrove']);
  draw(s, [
    [IMG, 11.01,1.22,.25,.25, IC_SLATE],
  ]);
}

// Slide 16 — Headline, paragraph, photo row and an icon note
function slide16(s) {
  chrome(s, 16);
  tx(s, {x:1.21, y:.79, w:5.28, h:1.92, sz:36, c:BLACK, f:HEAD, af:1}, ['Create Persuade The ', ['Ultimate Pencil Solution', {c:BLUE}], ' Maximum']);
  tx(s, {x:7.71, y:1.18, w:4.33, h:1.13, sz:12, c:INK, ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor, aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus donec.']);
  tx(s, {x:5.75, y:5.6, w:3.68, h:.6, sz:12, c:INK, ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget']);
  draw(s, [
    [RECT, 4.8,5.49,.81,.81, BLUE, {sd:1}],
    [IMG, 5.01,5.7,.39,.39, IC_WHITE],
  ]);
}

// Slide 17 — Break slide — banner image with a caption strip
function slide17(s) {
  chrome(s, 17);
  draw(s, [
    [RECT, 1.43,2.56,10.87,2.56, BLUE],
  ]);
  tx(s, {x:1.29, y:5.3, w:11.15, h:.6, sz:12, c:INK, ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla cons quis.']);
  tx(s, {x:3.45, y:1.59, w:8.75, h:.67, sz:34, c:BLACK, f:HEAD, nw:1, af:1}, ['Creative ', ['Pencil Strategies', {c:BLUE}], ' That Sell']);
  tx(s, {x:1.43, y:1.98, w:1.82, h:.58, sz:14, c:BLUE, f:HEAD, a:'center', v:1, f2:WHITE, sd:1}, ['Break Slide']);
}

// Slide 18 — Looping ribbon pencil beside four stat tiles
function slide18(s) {
  chrome(s, 18);
  draw(s, [
    [P18A, 1.9,.57,.12,5.43, SLATE_D],
    [P18A, 2.19,.57,.12,5.43, SLATE_L],
    [RECT, -.61,3.19,5.43,.2, SLATE, {r:270.2}],
    [P18B, 1.9,5.91,.41,.83, GREY],
    [P18C, 2.03,6.42,.15,.31, CHAR],
    [P18D, 1.09,3.23,4.42,1.71, PINK],
    [P18E, 1.39,3.56,3.58,1.06, BLUE],
    [P18F, 0,1.21,3.43,2.31, PINK],
    [P18G, 0,1.51,3.12,1.71, TEAL],
    [P18H, 0,1.83,2.81,1.05, GOLD],
    [P18I, -.01,.91,3.71,2.92, BLUE],
    [P18J, .81,2.93,5.27,2.32, TEAL],
    [P18K, .52,2.63,6.12,2.93, GOLD],
    ['roundRect', 9.74,3.31,2.27,1.38, WHITE, {sd:1,rr:.09}],
    ['roundRect', 7.4,4.75,2.27,1.38, WHITE, {sd:1,rr:.09}],
  ]);
  tx(s, {x:7.26, y:1.37, w:5.01, h:1.45, sz:40, c:INK, f:HEAD, af:1}, ['Pencil ', ['Solutions for', {c:BLUE}], ' Maximum']);
  draw(s, [
    ['roundRect', 9.74,4.75,2.27,1.38, PINK, {sd:1,rr:.09}],
    ['roundRect', 7.4,3.31,2.27,1.38, GOLD, {sd:1,rr:.09}],
  ]);
  tx(s, {x:7.67, y:3.52, w:1.73, h:.71, sz:40, c:WHITE, f:HEAD, a:'center', ls:.9, af:1}, [['375+', {b:1}]]);
  tx(s, {x:7.67, y:4.21, w:1.73, h:.27, sz:11, c:WHITE, a:'center', ls:.9, af:1}, ['Total Data In 2020']);
  tx(s, {x:10.01, y:3.52, w:1.73, h:.71, sz:40, c:TEAL, f:HEAD, a:'center', ls:.9, af:1}, [['234+', {b:1}]]);
  tx(s, {x:10.01, y:4.21, w:1.73, h:.27, sz:11, c:TEAL, a:'center', ls:.9, af:1}, ['Total Data In 2021']);
  tx(s, {x:10.01, y:4.96, w:1.73, h:.71, sz:40, c:WHITE, f:HEAD, a:'center', ls:.9, af:1}, [['510+', {b:1}]]);
  tx(s, {x:10.01, y:5.65, w:1.73, h:.27, sz:11, c:WHITE, a:'center', ls:.9, af:1}, ['Total Data In 2023']);
  tx(s, {x:7.67, y:4.96, w:1.73, h:.71, sz:40, c:BLUE, f:HEAD, a:'center', ls:.9, af:1}, [['408+', {b:1}]]);
  tx(s, {x:7.67, y:5.65, w:1.73, h:.27, sz:11, c:SLATE, a:'center', ls:.9, af:1}, ['Total Data In 2022']);
}

// Slide 19 — Pencil "tree" of leaves with four corner labels
function slide19(s) {
  chrome(s, 19);
  draw(s, [
    [P19A, 6.83,1.17,.9,.74, SLATE_D, {r:157,fv:1}],
    [P19B, 8.39,1.98,.43,.54, PURPLE_L, {r:27.7}],
    [P19C, 6.94,1.38,1.56,1.16, SLATE, {r:330.6}],
    [P19B, 6.55,1.98,.43,.54, BLUE_L, {r:346}],
    [P19D, 4.8,.84,1.41,1.05, TEAL, {r:207.2,fv:1}],
    [P19E, 7.07,2.89,1.6,1.07, GREY],
    [P19F, 6.24,1.88,.15,.97, GREY],
    [P19G, 5.79,3.67,.21,2.06, GREY],
    [P19H, 7.37,4.07,.18,1.67, GREY],
    [P19I, 7.02,3.6,.14,1.82, GREY],
    [P19J, 5.88,5.3,1.61,1.35, GREY],
    [P19K, 4.3,2.25,1.79,1.61, GOLD],
    [P19L, 7.29,2.29,1.75,1.28, PINK],
    [P19M, 6.22,3.18,.11,2.24, GREY],
    [P19N, 6.19,2.73,1.08,.87, PINK_L],
    [P19O, 5.39,1.96,.55,.72, GOLD_L],
    [P19P, 7.39,3.63,.89,.83, BLUE],
    [P19Q, 7.03,2.03,.56,.7, SLATE_L],
    [P19R, 5.89,1.25,.79,1.01, TEAL_L],
    [P19S, 4.55,1.82,.79,.64, GOLD_D],
    [P19T, 6.19,3.65,1.01,.75, PINK_D, {r:209.4,fv:1}],
    [P19B, 7.21,3.85,.43,.54, BLUE_D, {r:338.1}],
  ]);
  tx(s, {x:9.66, y:1.89, w:2.62, h:.4, sz:18, c:SLATE_T, f:HEAD, af:1}, [['Title Three', {b:1}]]);
  tx(s, {x:9.66, y:2.28, w:2.99, h:.69, sz:14, c:BLACK, ls:1.3, af:1}, ['Far far', T4]);
  tx(s, {x:9.66, y:5.37, w:2.62, h:.4, sz:18, c:SLATE_T, f:HEAD, af:1}, [['Title Four', {b:1}]]);
  tx(s, {x:9.66, y:5.76, w:2.99, h:.69, sz:14, c:BLACK, ls:1.3, af:1}, ['Far far', T4]);
  tx(s, {x:1.14, y:1.89, w:2.62, h:.4, sz:18, c:SLATE_T, f:HEAD, a:'right', af:1}, [['Title One', {b:1}]]);
  tx(s, {x:.77, y:2.28, w:2.99, h:.69, sz:14, c:BLACK, a:'right', ls:1.3, af:1}, ['Far far', T4]);
  tx(s, {x:1.14, y:5.37, w:2.62, h:.4, sz:18, c:SLATE_T, f:HEAD, a:'right', af:1}, [['Title Two', {b:1}]]);
  tx(s, {x:.77, y:5.76, w:2.99, h:.69, sz:14, c:BLACK, a:'right', ls:1.3, af:1}, ['Far far', T4]);
  draw(s, [
    [P19U, 3.25,1.28,.29,.27, WHITE],
    [P19V, 3.26,4.76,.27,.27, WHITE],
    [P19W, 9.93,1.28,.27,.27, WHITE],
    [P19X, 9.93,4.76,.27,.27, WHITE],
    ['flowChartConnector', 9.73,1.28,.51,.51, BLUE, {fv:1,sd:1}],
    [IMG, 9.86,1.41,.25,.25, IC_WHITE],
    ['flowChartConnector', 9.73,4.75,.51,.51, PINK, {fv:1,sd:1}],
    [IMG, 9.88,4.88,.26,.26, IC_WHITE],
    ['flowChartConnector', 3.19,1.28,.51,.51, TEAL, {fv:1,sd:1}],
    [IMG, 3.35,1.41,.24,.24, IC_WHITE],
    ['flowChartConnector', 3.19,4.75,.51,.51, GOLD, {fv:1,sd:1}],
    [IMG, 3.34,4.88,.22,.22, IC_WHITE],
  ]);
}

// Slide 20 — Diagonal pencil with four chevron markers
function slide20(s) {
  chrome(s, 20);
  draw(s, [
    [P14B, 5.42,1.93,.71,.38, BLUE],
    [P20A, 5.06,1.93,1.07,.77, BLUE_L],
    [P14B, 3.89,3.47,.71,.39, TEAL],
    [P20A, 3.53,3.47,1.07,.77, TEAL_L],
    [P14B, 7.19,4.79,.72,.39, GOLD],
    [P20B, 7.19,4.79,1.08,.77, GOLD_L],
    [P14B, 8.72,3.26,.72,.38, PINK],
    [P20B, 8.72,3.26,1.08,.77, PINK_L],
    [P6A, 4.71,5.07,.35,.29, SNOW, {r:45}],
    [P6B, 5.03,5.38,.35,.29, GREY3, {r:45}],
    [P6C, 4.91,5.22,.27,.29, GREY, {r:45}],
    [P6D, 4.88,5.49,.14,.15, CHAR2, {r:45}],
    [P6E, 4.78,5.39,.14,.15, CHAR4, {r:45}],
    [P6F, 4.84,5.44,.1,.15, CHAR3, {r:45}],
    [P6G, 5.11,4.35,.8,1.11, GOLD_L, {r:45}],
    [P6H, 5.57,4.53,.27,1.11, GOLD_D, {r:45}],
    [P6I, 5.38,4.35,.27,1.11, GOLD, {r:45}],
    [RECT, 5.88,3.6,.8,1.09, TEAL_L, {r:45}],
    [RECT, 6.33,3.78,.27,1.09, TEAL_D, {r:45}],
    [RECT, 6.14,3.6,.27,1.09, TEAL, {r:45}],
    [RECT, 6.65,2.83,.8,1.09, PINK_L, {r:45}],
    [RECT, 7.1,3.01,.27,1.09, PINK_D, {r:45}],
    [RECT, 6.91,2.83,.27,1.09, PINK, {r:45}],
    [RECT, 7.41,2.06,.8,1.09, BLUE_L, {r:45}],
    [RECT, 7.86,2.25,.27,1.09, BLUE_D, {r:45}],
    [RECT, 7.68,2.06,.27,1.09, BLUE, {r:45}],
  ]);
  tx(s, {x:6.46, y:.55, w:1.63, h:.53, sz:18, c:BLUE, f:HEAD, a:'right', ls:1.5, af:1}, ['Creative A']);
  tx(s, {x:5.51, y:1.1, w:2.57, h:.6, sz:12, c:BLACK, a:'right', ls:1.3, sb:12, af:1}, ['Far from the countries Vokalia and Consonantia, there live']);
  tx(s, {x:9.14, y:4.2, w:3.34, h:.53, sz:18, c:PINK, f:HEAD, ls:1.5, af:1}, ['Creative B']);
  tx(s, {x:9.14, y:4.75, w:2.77, h:.6, sz:12, c:BLACK, ls:1.3, sb:12, af:1}, ['Far from the countries Vokalia and Consonantia, there live']);
  tx(s, {x:.75, y:4.36, w:3.34, h:.55, sz:18, c:TEAL, f:HEAD, a:'right', ls:1.5, af:1}, ['Creative C']);
  tx(s, {x:1.32, y:4.92, w:2.77, h:.61, sz:12, c:BLACK, a:'right', ls:1.3, sb:12, af:1}, ['Far from the countries Vokalia and Consonantia, there live']);
  tx(s, {x:5.66, y:5.67, w:1.68, h:.53, sz:18, c:GOLD, f:HEAD, ls:1.5, af:1}, ['Creative D']);
  tx(s, {x:5.66, y:6.22, w:2.6, h:.6, sz:12, c:BLACK, ls:1.3, sb:12, af:1}, ['Far from the countries Vokalia and Consonantia, there live']);
  draw(s, [
    ['ellipse', 3.6,3.93,.47,.47, WHITE, {sd:1}],
    [IMG, 3.73,4.04,.24,.24, IC_TEAL],
    ['ellipse', 5.73,5.27,.47,.47, WHITE, {sd:1}],
    [IMG, 5.84,5.38,.24,.24, IC_GOLD],
    ['ellipse', 9.24,3.75,.47,.47, WHITE, {sd:1}],
    [IMG, 9.37,3.85,.26,.26, IC_PINK],
    ['ellipse', 7.55,1.76,.47,.47, WHITE, {sd:1}],
    [IMG, 7.66,1.88,.24,.24, IC_BLUE],
  ]);
}

// Slide 21 — Four interlocking puzzle pieces forming a pencil
function slide21(s) {
  s.background = { color: BLUE };
  chrome(s, 21, true);
  draw(s, [
    [P4D, 9.88,2.66,2.12,2.18, GREY, {r:180.4,fv:1}],
    [P4F, 11.38,3.41,.61,.71, BLACK, {r:180.4,fv:1}],
    [P21A, 5.7,2.67,2.87,2.84, PINK, {r:270,lc:PINK_W,lw:1.5}],
    [P21B, 1.01,2.98,2.87,2.21, PINK, {r:270,lc:GOLD_W,lw:2}],
    [P21C, 3.22,1.66,2.87,3.5, TEAL, {r:90,fv:1,lc:TEAL_W,lw:2}],
    [P21B, 7.57,2.32,2.87,2.21, TEAL, {r:90,lc:BLUE_P,lw:2}],
  ]);
  tx(s, {x:1.34, y:5.72, w:1.11, h:.34, sz:14, c:PINK, f:HEAD, nw:1, af:1}, ['Title One']);
  tx(s, {x:1.34, y:6.06, w:1.94, h:.6, sz:12, c:WHITE, ls:1.3, sb:12, af:1}, ['Separated they live in Bookmarksgrove']);
  tx(s, {x:3.55, y:.84, w:1.12, h:.34, sz:14, c:TEAL, f:HEAD, nw:1, af:1}, ['Title Two']);
  tx(s, {x:3.55, y:1.17, w:1.94, h:.6, sz:12, c:WHITE, ls:1.3, af:1}, ['Separated they live in Bookmarksgrove']);
  tx(s, {x:5.71, y:5.72, w:1.26, h:.34, sz:14, c:PINK, f:HEAD, nw:1, af:1}, ['Title Three']);
  tx(s, {x:5.71, y:6.06, w:1.94, h:.6, sz:12, c:WHITE, ls:1.3, af:1}, ['Separated they live in Bookmarksgrove']);
  tx(s, {x:7.89, y:.84, w:1.14, h:.34, sz:14, c:TEAL, f:HEAD, nw:1, af:1}, ['Title Four']);
  tx(s, {x:7.89, y:1.17, w:1.94, h:.6, sz:12, c:WHITE, ls:1.3, af:1}, ['Separated they live in Bookmarksgrove']);
  draw(s, [
    [IMG, 6.82,3.33,.81,.81, IC_WHITE],
    [IMG, 4.33,3.41,.68,.68, IC_WHITE],
    [IMG, 8.92,3.41,.66,.66, IC_WHITE],
    [IMG, 1.82,3.44,.64,.64, IC_WHITE],
  ]);
}

// Slide 22 — Circular pencil ring with four outer labels
function slide22(s) {
  chrome(s, 22);
  tx(s, {x:4.82, y:2.94, w:3.69, h:1.45, sz:40, c:BLUE, f:HEAD, a:'center', af:1}, [['Pencil ', {c:INK, b:1}], ['Infographic', {b:1}]]);
  draw(s, [
    [P22A, 5.68,5.72,.62,.72, PURPLE],
    [RECT, 5.36,5.74,.65,.42, GREY, {r:295}],
    [RECT, 5.36,5.94,.65,.03, GREY5, {r:294.9}],
    [RECT, 5.19,5.86,.65,.03, GREY5, {r:295}],
    [RECT, 5.53,6.01,.65,.03, GREY5, {r:295}],
    [P22B, 5.68,6.26,.03,.03, GREY3, {r:295}],
    [P22B, 5.92,5.73,.03,.03, GREY3, {r:295}],
    [P22B, 5.8,6,.03,.03, GREY3, {r:295}],
    [P22B, 5.74,6.13,.03,.03, GREY3, {r:295}],
    [P22B, 5.71,6.2,.03,.03, GREY3, {r:295}],
    [P22B, 5.77,6.06,.03,.03, GREY3, {r:295}],
    [P22B, 5.86,5.87,.03,.03, GREY3, {r:295}],
    [P22B, 5.83,5.93,.03,.03, GREY3, {r:295}],
    [P22B, 5.89,5.8,.03,.03, GREY3, {r:295}],
    [P22B, 5.42,6.14,.03,.03, GREY3, {r:295}],
    [P22B, 5.67,5.61,.03,.03, GREY3, {r:295}],
    [P22B, 5.54,5.88,.03,.03, GREY3, {r:295}],
    [P22B, 5.48,6.01,.03,.03, GREY3, {r:295}],
    [P22B, 5.45,6.08,.03,.03, GREY3, {r:295}],
    [P22B, 5.51,5.94,.03,.03, GREY3, {r:295}],
    [P22B, 5.6,5.74,.03,.03, GREY3, {r:295}],
    [P22B, 5.57,5.81,.03,.03, GREY3, {r:295}],
    [P22B, 5.63,5.68,.03,.03, GREY3, {r:295}],
    [P22C, 3.95,1.06,5.43,5.32, BLUE_L],
    [P22D, 4.15,1.26,5.03,4.93, BLUE],
    [P22E, 4.4,1.51,4.54,4.44, BLUE_D],
    [P22F, 6.9,5.75,.56,.63, GREY],
    [P22G, 6.7,6.03,.29,.27, CHAR],
  ]);
  tx(s, {x:9.54, y:5.04, w:2.35, h:.4, sz:20, c:SLATE, f:HEAD, nw:1}, ['Title Four']);
  tx(s, {x:9.54, y:5.51, w:2.87, h:.81, sz:14, c:BLACK, af:1}, ['Far far', T5, 'Vokalia and.']);
  draw(s, [
    ['flowChartConnector', 8.41,5.09,.91,.91, WHITE, {sd:1}],
  ]);
  tx(s, {x:9.54, y:1.18, w:2.35, h:.4, sz:20, c:PINK, f:HEAD, nw:1}, ['Title Three']);
  tx(s, {x:9.54, y:1.65, w:2.87, h:.81, sz:14, c:BLACK, af:1}, ['Far far', T5, 'Vokalia and.']);
  draw(s, [
    ['flowChartConnector', 8.41,1.23,.91,.91, WHITE, {sd:1}],
  ]);
  tx(s, {x:1.44, y:1.18, w:2.35, h:.4, sz:20, c:GOLD, f:HEAD, a:'right', nw:1}, ['Title Two']);
  tx(s, {x:.93, y:1.65, w:2.87, h:.81, sz:14, c:BLACK, a:'right', af:1}, ['Far far', T5, 'Vokalia and.']);
  draw(s, [
    ['flowChartConnector', 4.04,1.23,.91,.91, WHITE, {sd:1}],
  ]);
  tx(s, {x:1.44, y:5.04, w:2.35, h:.4, sz:20, c:TEAL, f:HEAD, a:'right', nw:1}, ['Title One']);
  tx(s, {x:.93, y:5.51, w:2.87, h:.81, sz:14, c:BLACK, a:'right', af:1}, ['Far far', T5, 'Vokalia and.']);
  draw(s, [
    ['flowChartConnector', 4.04,5.09,.91,.91, WHITE, {sd:1}],
    [IMG, 8.67,1.47,.43,.43, IC_PINK],
    [IMG, 8.68,5.35,.38,.38, IC_SLATE],
    [IMG, 4.31,5.36,.36,.36, IC_TEAL],
    [IMG, 4.31,1.5,.37,.37, IC_GOLD],
  ]);
}

// Slide 23 — Diagonal stacked pencil with four ribbon rows
function slide23(s) {
  chrome(s, 23);
  draw(s, [
    [P23A, 2.02,3.79,4.57,.99, SNOW, {tr:50}],
    [P23A, 2.94,1.9,4.54,1, SNOW, {tr:50}],
    [P14B, 2.63,1.9,.69,.5, PINK_D],
    [P20A, 2.29,1.9,1.03,.99, PINK],
    [P23A, 5.86,3.68,5.28,.99, SNOW, {r:180}],
    [P14B, 10.65,4.18,.69,.5, TEAL_D, {r:180}],
    [P20A, 10.65,3.68,1.03,.99, TEAL, {r:180}],
    [P14B, 1.66,3.78,.69,.5, GOLD_D],
    [P20A, 1.31,3.78,1.03,.99, GOLD],
    [P23A, 5.9,5.62,4.77,.99, SNOW, {r:180}],
    [P14B, 10.25,6.12,.59,.5, BLUE_D, {r:180}],
    [P20A, 10.25,5.62,.89,.99, BLUE, {r:180}],
    [RECT, 6.82,1.97,.91,1.23, PINK, {r:19.7}],
    [P4H, 6.41,1.74,.41,1.43, PINK, {r:19.7,tr:20}],
    [P23B, 7.65,2.19,.41,1.43, PINK_D, {r:19.7,tr:10}],
    [P23C, 6.15,3.16,1.74,.28, SNOW, {r:19.7}],
    [P4B, 6.59,3.16,.91,.11, GREY, {r:19.7}],
    [P23D, 7.35,3.37,.49,.28, GREY2, {r:19.7}],
    [RECT, 6.42,3.23,.8,1.08, TEAL, {r:19.7}],
    [P4H, 6.06,3.02,.37,1.26, TEAL, {r:19.7,tr:20}],
    [P23B, 7.16,3.42,.37,1.26, TEAL_D, {r:19.7,tr:10}],
    [P23C, 5.83,4.27,1.53,.24, SNOW, {r:19.7}],
    [P23E, 6.22,4.27,.8,.09, GREY, {r:19.7}],
    [P23D, 6.89,4.45,.43,.24, GREY2, {r:19.7}],
    [RECT, 6.09,4.33,.69,.94, GOLD, {r:19.7}],
    [P4H, 5.78,4.16,.32,1.09, GOLD, {r:19.7,tr:20}],
    [P23B, 6.72,4.5,.32,1.09, GOLD_D, {r:19.7,tr:10}],
    [P23C, 5.58,5.23,1.33,.2, SNOW, {r:19.7}],
    [P4B, 5.92,5.24,.69,.07, GREY, {r:19.7}],
    [P23D, 6.5,5.39,.38,.2, GREY2, {r:19.7}],
    [RECT, 5.79,5.28,.61,.83, BLUE_D, {r:19.7}],
    [P4H, 5.51,5.12,.28,.96, BLUE, {r:19.7,tr:20}],
    [P23B, 6.35,5.42,.28,.96, BLUE_XD, {r:19.7,tr:10}],
    [P23F, 5.31,6.07,1.17,.35, SKIN_D, {r:19.7}],
    [P14B, 5.53,6.06,.61,.72, GREY, {r:19.7}],
    [P23G, 5.27,5.96,.58,.72, SNOW, {r:19.7}],
    [P23H, 5.82,6.16,.58,.72, GREY3, {r:19.7}],
    [P23I, 5.54,6.48,.42,.29, SLATE_T, {r:19.7}],
  ]);
  tx(s, {x:3.99, y:1.97, w:2.08, h:.36, sz:14, c:PINK, f:HEAD, a:'right', ls:1.2, af:1}, ['Title One']);
  tx(s, {x:3.4, y:2.26, w:2.68, h:.57, sz:12, c:BLACK, a:'right', ls:1.2, af:1}, ['Far far', T6]);
  tx(s, {x:7.81, y:3.74, w:2.08, h:.36, sz:14, c:TEAL, f:HEAD, ls:1.2, af:1}, ['Title Two']);
  tx(s, {x:7.81, y:4.04, w:2.52, h:.57, sz:12, c:BLACK, ls:1.2, af:1}, ['Far far', T6]);
  tx(s, {x:3.52, y:3.85, w:2.08, h:.36, sz:14, c:GOLD, f:HEAD, a:'right', ls:1.2, af:1}, ['Title Three']);
  tx(s, {x:2.78, y:4.15, w:2.81, h:.57, sz:12, c:BLACK, a:'right', ls:1.2, af:1}, ['Far far', T6]);
  tx(s, {x:6.89, y:5.7, w:2.08, h:.36, sz:14, c:BLUE, f:HEAD, ls:1.2, af:1}, ['Title Four']);
  tx(s, {x:6.89, y:5.99, w:2.81, h:.57, sz:12, c:BLACK, ls:1.2, af:1}, ['Far far', T6]);
  tx(s, {x:2.61, y:.5, w:8.12, h:.77, sz:40, c:BLACK, f:HEAD, a:'center', ls:.9, sb:10}, ['The Pencil ', ['Effect', {c:BLUE}]]);
}

// Slide 24 — Headline, note row, stat strip and side photo
function slide24(s) {
  chrome(s, 24);
  tx(s, {x:1.75, y:1.35, w:5.63, h:2.12, sz:40, c:BLACK, f:HEAD, af:1}, ['Sketching ', ['Success Innovative Pencil ', {c:BLUE}], 'Solutions Impact']);
  draw(s, [
    [RECT, 7.63,5.31,2.95,.83, WHITE, {sd:1}],
    [RECT, 2.95,5.31,4.52,.83, BLUE],
    [RECT, 1.89,5.31,.9,.83, WHITE, {sd:1}],
  ]);
  tx(s, {x:1.78, y:3.81, w:1.29, h:.37, sz:16, c:BLACK, f:HEAD, nw:1, af:1}, ['Your Text']);
  tx(s, {x:1.78, y:4.18, w:5.69, h:.6, sz:12, c:INK, ls:1.3, af:1}, ['Lorem ipsum dolor  amet, consectetuer adipiscing. Aenean commodo ligula eget dolor. aenean Cum sociis natoque penatibus et magnis dis']);
  draw(s, [
    [IMG, 2.09,5.48,.5,.5, IC_BLUE],
  ]);
  tx(s, {x:7.89, y:5.46, w:1.7, h:.57, sz:28, c:BLUE, f:HEAD, nw:1, af:1}, ['13.478+']);
  tx(s, {x:9.5, y:5.62, w:.83, h:.34, sz:14, c:GREY4, nw:1, af:1}, ['People']);
  draw(s, [
    [IMG, 3.09,5.45,.23,.23, IC_WHITE],
  ]);
  tx(s, {x:3.3, y:5.42, w:3.98, h:.29, sz:11, c:WHITE, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing. ']);
  draw(s, [
    [IMG, 3.09,5.78,.23,.23, IC_WHITE],
  ]);
  tx(s, {x:3.3, y:5.75, w:2.78, h:.29, sz:11, c:WHITE, af:1}, ['Aenean commodo ligula eget dolor.']);
}

// Slide 25 — Closing slide — Thank You over a ghost pencil
function slide25(s) {
  chrome(s, 25);
  draw(s, [
    [P1A, 12.02,5.01,.29,2.49, SNOW, {r:180,tr:35}],
    [P1B, 7.93,5.01,.29,2.49, SNOW, {r:180,tr:35}],
    [P1C, 8.86,5.91,.3,1.59, SNOW, {r:180,tr:35}],
    [P1D, 7.92,2.15,4.41,3.7, SNOW, {r:180,tr:35}],
    [P1E, 11.08,5.91,.29,1.59, SNOW, {r:180,tr:35}],
  ]);
  tx(s, {x:1, y:4.12, w:5.55, h:1.31, sz:72, c:BLUE, f:HEAD, nw:1, af:1}, ['Thank ', ['You', {c:BLACK}]]);
  tx(s, {x:1.2, y:5.25, w:5.3, h:.6, sz:12, c:INK, ls:1.3, af:1}, ['Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula dolor aenean massa far far away.']);
}

// ── build ───────────────────────────────────────────────────────────────────
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
pptx.layout = 'DECK';
pptx.theme = { headFontFace: HEAD, bodyFontFace: BODY };

const SLIDES = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25];
SLIDES.forEach(build => build(pptx.addSlide()));

pptx.writeFile({ fileName: path.join(__dirname, '083f04c6-cddf-4372-ad53-2a7bd478f4e8_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
