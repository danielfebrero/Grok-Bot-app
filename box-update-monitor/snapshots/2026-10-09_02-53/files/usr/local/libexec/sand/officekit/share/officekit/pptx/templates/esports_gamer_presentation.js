/**
 * Recreates "Gamer — Esport Presentation Template" (20 slides, 13.333 x 7.5 in)
 * with pptxgenjs.  Photographs in the source deck are replaced by flat grey
 * placeholder shapes that keep the original silhouette, position and size.
 */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */
const BG = '1B1A19';   // theme accent2 – deck background
const G = '00FF93';   // theme accent1 – signature green
const G75 = '00BF6E';   // accent1 lumMod 75%
const G60 = '66FFBE';   // accent1 lumMod 60% / lumOff 40%
const G50 = '00804A';   // accent1 lumMod 50%
const G40 = '99FFD4';   // accent1 lumMod 40% / lumOff 60%
const W = 'FFFFFF';
const INK = '262626';   // tx1 lumMod 85% / lumOff 15%
const IMG = 'CCCCCC';   // photo placeholder
const IMGTX = '8C8C8C';

const DISPLAY = 'Russo One';
const BODY = 'Work Sans';

const NOLINE = { type: 'none' };
const SHADOW = { type: 'outer', color: '000000', opacity: 0.2, blur: 50, offset: 20, angle: 45 };
const SOFTSHADOW = { type: 'outer', color: '000000', opacity: 0.2, blur: 13, offset: 3, angle: 45 };

/* --------------------------------------------------------- geometry helpers */
const M = (x, y) => ({ x, y, moveTo: true });
const L = (x, y) => ({ x, y });
const B = (x1, y1, x2, y2, x, y) => ({ x, y, curve: { type: 'cubic', x1, y1, x2, y2 } });
const Z = { close: true };

/** Hexagon outline inscribed in the box (x0,y0)-(x1,y1); `ins` = flat-top inset. */
const hexPath = (x0, y0, x1, y1, ins = 0.2155) => {
  const w = x1 - x0, h = y1 - y0, my = y0 + h / 2;
  return [M(x0 + ins * w, y0), L(x1 - ins * w, y0), L(x1, my), L(x1 - ins * w, y1), L(x0 + ins * w, y1), L(x0, my), Z];
};
/** Hexagonal ring: full-box hexagon minus a hexagon inset by `b`. */
const hexRing = (b) => [...hexPath(0, 0, 1, 1), ...hexPath(b, b, 1 - b, 1 - b)];

/** Scales a 0..1 normalised outline to `w` x `h` inches for pptxgenjs. */
const scale = (pts, w, h) => pts.map((p) => (p.close ? p : p.curve
  ? { x: p.x * w, y: p.y * h, curve: { type: 'cubic', x1: p.curve.x1 * w, y1: p.curve.y1 * h, x2: p.curve.x2 * w, y2: p.curve.y2 * h } }
  : { x: p.x * w, y: p.y * h, moveTo: p.moveTo }));

/** Adds a free-form shape whose outline is given in 0..1 box coordinates. */
function freeform(s, [x, y, w, h], points, opts = {}) {
  s.addShape('custGeom', { x, y, w, h, line: NOLINE, ...opts, points: scale(points, w, h) });
}

/* ------------------------------------------------------- icon outlines (0..1) */
const ICON = {
  logo: [M(0.34,0.49),B(0.34,0.6,0.41,0.69,0.5,0.69),B(0.59,0.69,0.66,0.6,0.66,0.49),B(0.66,0.38,0.59,0.29,0.5,0.29),B(0.41,0.29,0.34,0.38,0.34,0.49),Z,M(0.23,0.42),B(0.25,0.27,0.37,0.16,0.5,0.16),B(0.57,0.16,0.64,0.18,0.69,0.24),B(0.71,0.27,0.77,0.27,0.78,0.24),B(0.8,0.22,0.8,0.18,0.78,0.14),B(0.71,0.04,0.61,0,0.5,0),B(0.32,0,0.16,0.14,0.12,0.36),L(0,0.36),L(0,0.49),L(0.16,0.49),B(0.22,0.49,0.22,0.45,0.23,0.42),Z,M(0.84,0.49),B(0.78,0.49,0.78,0.56,0.77,0.58),B(0.75,0.73,0.64,0.84,0.5,0.84),B(0.43,0.84,0.36,0.8,0.3,0.73),B(0.29,0.71,0.23,0.71,0.22,0.73),B(0.2,0.78,0.2,0.82,0.22,0.84),B(0.29,0.93,0.39,1,0.5,1),B(0.68,1,0.84,0.84,0.88,0.64),L(1,0.64),L(1,0.49),L(0.84,0.49),Z],
  trophy: [M(0.56,0.73),B(0.56,0.65,0.62,0.61,0.72,0.56),B(0.84,0.48,1,0.38,1,0.15),B(1,0.13,0.98,0.11,0.96,0.11),L(0.78,0.11),B(0.74,0.06,0.66,0,0.5,0),B(0.34,0,0.26,0.06,0.22,0.11),L(0.04,0.11),B(0.02,0.11,0,0.13,0,0.15),B(0,0.38,0.14,0.48,0.28,0.56),B(0.38,0.61,0.44,0.65,0.44,0.73),L(0.44,0.81),B(0.32,0.82,0.24,0.86,0.24,0.9),B(0.24,0.96,0.36,1,0.5,1),B(0.64,1,0.74,0.96,0.74,0.9),B(0.74,0.86,0.68,0.82,0.56,0.81),L(0.56,0.73),Z,M(0.72,0.46),B(0.76,0.4,0.78,0.31,0.78,0.19),L(0.92,0.19),B(0.9,0.33,0.82,0.4,0.72,0.46),Z,M(0.5,0.08),B(0.66,0.08,0.72,0.13,0.72,0.15),B(0.72,0.17,0.66,0.23,0.5,0.25),B(0.34,0.23,0.28,0.17,0.28,0.15),B(0.28,0.13,0.34,0.08,0.5,0.08),Z,M(0.08,0.19),L(0.22,0.19),B(0.22,0.31,0.24,0.4,0.28,0.46),B(0.18,0.4,0.08,0.33,0.08,0.19),Z],
  rocket: [M(0.61,0.65),B(0.61,0.65,1,0.37,0.96,0.06),L(0.96,0.04),L(0.94,0.04),B(0.63,0,0.35,0.39,0.35,0.39),B(0.12,0.34,0.14,0.4,0.02,0.65),B(0,0.71,0.04,0.71,0.08,0.71),L(0.19,0.67),L(0.33,0.81),L(0.29,0.92),B(0.27,0.96,0.29,1,0.33,0.98),B(0.59,0.86,0.65,0.88,0.61,0.65),Z,M(0.67,0.32),B(0.63,0.29,0.63,0.25,0.67,0.21),B(0.71,0.17,0.76,0.17,0.78,0.21),B(0.82,0.25,0.82,0.29,0.78,0.32),B(0.76,0.37,0.71,0.37,0.67,0.32),Z],
  arcBg: [M(0,1),L(0.179,1),B(0.176,0.666,0.303,0.384,0.471,0.356),B(0.654,0.324,0.814,0.602,0.823,0.968),L(1,0.968),B(0.993,0.427,0.769,-0.004,0.497,0),B(0.221,0.004,-0.001,0.451,0,1),Z],
  arcDn: [M(0.871,0),L(0.743,0.213),L(0.788,0.213),B(0.787,0.502,0.641,0.736,0.464,0.721),B(0.382,0.715,0.31,0.655,0.257,0.564),B(0.204,0.472,0.172,0.349,0.171,0.214),L(0,0.212),B(0,0.64,0.207,0.99,0.466,1),B(0.735,1.01,0.955,0.654,0.958,0.213),L(1,0.213),L(0.871,0),Z],
  arcUp: [M(0.871,1),L(0.743,0.787),L(0.788,0.787),B(0.787,0.498,0.641,0.264,0.464,0.279),B(0.382,0.285,0.31,0.345,0.257,0.436),B(0.204,0.528,0.172,0.651,0.171,0.786),L(0,0.788),B(0,0.36,0.207,0.01,0.466,0),B(0.735,-0.01,0.955,0.346,0.958,0.787),L(1,0.787),L(0.871,1),Z],
  check: [M(0.5,0),B(0.37,0,0.24,0.05,0.15,0.15),B(-0.05,0.34,-0.05,0.66,0.15,0.85),B(0.34,1.05,0.66,1.05,0.85,0.85),B(1.05,0.66,1.05,0.34,0.85,0.15),B(0.76,0.05,0.63,0,0.5,0),Z,M(0.74,0.29),L(0.79,0.37),L(0.41,0.7),L(0.21,0.46),L(0.28,0.43),L(0.43,0.59),L(0.74,0.29),Z],
  picture: [M(1,0.9),B(1,0.95,0.97,1,0.92,1),L(0.1,1),B(0.04,1,0,0.95,0,0.9),L(0,0.1),B(0,0.03,0.04,0,0.1,0),L(0.92,0),B(0.97,0,1,0.03,1,0.1),L(1,0.9),Z,M(0.1,0.07),L(0.07,0.1),L(0.07,0.9),L(0.1,0.91),L(0.92,0.91),L(0.93,0.9),L(0.93,0.1),B(0.93,0.09,0.93,0.07,0.92,0.07),L(0.1,0.07),Z,M(0.23,0.41),B(0.18,0.41,0.14,0.36,0.14,0.29),B(0.14,0.22,0.18,0.16,0.23,0.16),B(0.29,0.16,0.34,0.22,0.34,0.29),B(0.34,0.36,0.29,0.41,0.23,0.41),Z,M(0.88,0.83),L(0.14,0.83),L(0.14,0.71),L(0.3,0.5),L(0.38,0.6),L(0.66,0.26),L(0.88,0.53),L(0.88,0.83),Z],
  people: [M(0.18,0.57),L(0.11,0.57),B(0.05,0.57,0,0.54,0,0.49),B(0,0.43,0,0.28,0.07,0.28),B(0.08,0.28,0.14,0.32,0.21,0.32),L(0.27,0.31),L(0.27,0.35),B(0.27,0.4,0.29,0.46,0.32,0.5),B(0.26,0.5,0.21,0.53,0.18,0.57),Z,M(0.21,0.28),B(0.14,0.28,0.07,0.22,0.07,0.13),B(0.07,0.06,0.14,0,0.21,0),B(0.27,0,0.34,0.06,0.34,0.13),B(0.34,0.22,0.27,0.28,0.21,0.28),Z,M(0.73,1),L(0.27,1),B(0.19,1,0.14,0.94,0.14,0.85),B(0.14,0.72,0.16,0.53,0.32,0.53),B(0.34,0.53,0.4,0.6,0.51,0.6),B(0.6,0.6,0.67,0.53,0.68,0.53),B(0.85,0.53,0.88,0.72,0.88,0.85),B(0.88,0.94,0.82,1,0.73,1),Z,M(0.51,0.57),B(0.4,0.57,0.3,0.47,0.3,0.35),B(0.3,0.24,0.4,0.13,0.51,0.13),B(0.62,0.13,0.7,0.24,0.7,0.35),B(0.7,0.47,0.62,0.57,0.51,0.57),Z,M(0.81,0.28),B(0.73,0.28,0.67,0.22,0.67,0.13),B(0.67,0.06,0.73,0,0.81,0),B(0.88,0,0.93,0.06,0.93,0.13),B(0.93,0.22,0.88,0.28,0.81,0.28),Z,M(0.9,0.57),L(0.84,0.57),B(0.79,0.53,0.75,0.5,0.7,0.5),B(0.73,0.46,0.74,0.4,0.74,0.35),L(0.74,0.31),L(0.81,0.32),B(0.88,0.32,0.93,0.28,0.95,0.28),B(1,0.28,1,0.43,1,0.49),B(1,0.54,0.96,0.57,0.9,0.57),Z],
  envelope: [M(1,0.17),L(1,0.81),B(1,0.91,0.91,1,0.81,1),L(0.19,1),B(0.09,1,0,0.91,0,0.81),L(0,0.17),B(0,0.07,0.09,0,0.19,0),L(0.81,0),B(0.91,0,1,0.07,1,0.17),Z,M(0.78,0.24),L(0.22,0.24),B(0.19,0.24,0.16,0.28,0.16,0.31),B(0.16,0.34,0.21,0.38,0.22,0.4),B(0.29,0.43,0.34,0.47,0.4,0.52),L(0.5,0.55),B(0.53,0.55,0.57,0.53,0.59,0.52),L(0.78,0.4),B(0.79,0.38,0.83,0.34,0.83,0.31),B(0.83,0.28,0.81,0.24,0.78,0.24),Z,M(0.83,0.4),B(0.83,0.41,0.81,0.43,0.79,0.43),L(0.6,0.55),B(0.57,0.57,0.53,0.6,0.5,0.6),B(0.47,0.6,0.41,0.57,0.4,0.55),B(0.33,0.52,0.26,0.48,0.21,0.43),L(0.16,0.4),L(0.16,0.69),B(0.16,0.72,0.19,0.74,0.22,0.74),L(0.78,0.74),B(0.81,0.74,0.83,0.72,0.83,0.69),L(0.83,0.4),Z],
  doc1: [M(0.68,0),L(0.03,0.02),L(0,0.11),L(0,0.89),L(0.03,0.97),L(0.14,1),L(0.91,0.99),L(0.97,0.97),L(1,0.89),L(1,0.25),L(0.68,0),Z,M(0.57,0.34),L(0.57,0.06),L(0.93,0.34),L(0.57,0.34),Z],
  doc2: [M(0.14,0),L(0,0),L(0,0.89),L(0.04,0.97),L(0.14,1),L(1,1),L(1,0.89),L(0.14,0.89),L(0.14,0),Z],
  doc3: [M(0.14,0),L(0,0),L(0,0.89),L(0.04,0.98),L(0.14,1),L(1,1),L(1,0.89),L(0.14,0.89),L(0.14,0),Z],
  head: [M(0.504,0),L(0.594,0.006),L(0.669,0.024),L(0.752,0.06),L(0.839,0.124),L(0.894,0.146),L(0.939,0.19),L(0.959,0.233),L(0.968,0.271),L(0.968,0.368),L(0.95,0.453),L(0.941,0.477),L(0.976,0.482),L(0.992,0.497),L(1,0.525),L(0.99,0.585),L(0.955,0.666),L(0.921,0.707),L(0.899,0.715),L(0.885,0.758),L(0.861,0.804),L(0.807,0.871),L(0.76,0.911),L(0.704,0.947),L(0.641,0.975),L(0.571,0.994),L(0.5,1),L(0.428,0.994),L(0.358,0.976),L(0.293,0.947),L(0.239,0.912),L(0.191,0.871),L(0.153,0.827),L(0.125,0.782),L(0.106,0.736),L(0.101,0.715),L(0.079,0.707),L(0.045,0.666),L(0.005,0.568),L(0,0.525),L(0.012,0.491),L(0.037,0.478),L(0.059,0.477),L(0.036,0.404),L(0.031,0.352),L(0.035,0.302),L(0.052,0.252),L(0.08,0.203),L(0.117,0.161),L(0.184,0.105),L(0.292,0.044),L(0.373,0.015),L(0.436,0.004),L(0.504,0),Z],
  torso: [M(0.322,0),L(0.43,0.661),L(0.472,0.429),L(0.441,0.325),L(0.433,0.223),L(0.449,0.177),L(0.509,0.151),L(0.542,0.166),L(0.562,0.201),L(0.568,0.267),L(0.547,0.373),L(0.528,0.429),L(0.57,0.661),L(0.678,0.001),L(0.779,0.113),L(0.866,0.18),L(0.919,0.211),L(0.943,0.235),L(0.969,0.287),L(0.994,0.396),L(0.988,0.769),L(0.878,0.871),L(0.795,0.924),L(0.691,0.969),L(0.601,0.991),L(0.5,1),L(0.399,0.991),L(0.309,0.969),L(0.205,0.924),L(0.089,0.845),L(0.012,0.768),L(0.002,0.442),L(0.018,0.328),L(0.038,0.269),L(0.057,0.235),L(0.081,0.211),L(0.187,0.141),L(0.322,0),Z],
  money: [M(0.9,0.16),L(0.9,0.11),L(0.8,0.11),L(0.8,0.16),L(0.7,0.16),L(0.68,0),L(0.3,0.02),L(0.3,0.16),L(0.2,0.16),L(0.2,0.11),L(0.1,0.11),L(0.1,0.16),L(0.03,0.18),L(0,0.27),L(0,0.89),L(0.03,0.97),L(0.1,1),L(0.9,1),L(0.98,0.97),L(1,0.89),L(0.98,0.18),L(0.9,0.16),Z,M(0.4,0.11),L(0.61,0.11),L(0.61,0.16),L(0.4,0.16),L(0.4,0.11),Z,M(0.65,0.47),L(0.47,0.47),L(0.47,0.52),L(0.58,0.53),L(0.64,0.61),L(0.63,0.73),L(0.56,0.79),L(0.56,0.84),L(0.45,0.84),L(0.45,0.79),L(0.35,0.79),L(0.35,0.68),L(0.56,0.66),L(0.53,0.63),L(0.42,0.62),L(0.36,0.55),L(0.37,0.41),L(0.45,0.36),L(0.45,0.31),L(0.56,0.31),L(0.56,0.36),L(0.65,0.36),L(0.65,0.47),Z],
  safeA: [M(0.48,1),L(0.84,0.84),L(0.96,0.68),L(1,0.48),L(0.96,0.28),L(0.84,0.16),L(0.68,0.04),L(0.48,0),L(0.28,0.04),L(0.04,0.28),L(0,0.48),L(0.04,0.68),L(0.16,0.84),L(0.28,0.92),L(0.48,1),Z],
  safeB: [M(0.9,0),L(0.03,0.02),L(0,0.1),L(0.01,0.88),L(0.03,0.92),L(0.1,0.95),L(0.1,1),L(0.2,1),L(0.2,0.95),L(0.8,0.95),L(0.8,1),L(0.9,1),L(0.9,0.95),L(0.97,0.92),L(1,0.86),L(0.99,0.06),L(0.97,0.02),L(0.9,0),Z,M(0.85,0.81),L(0.2,0.81),L(0.2,0.66),L(0.1,0.66),L(0.1,0.57),L(0.2,0.57),L(0.2,0.38),L(0.1,0.38),L(0.1,0.29),L(0.2,0.29),L(0.2,0.14),L(0.85,0.14),L(0.85,0.81),Z],
  pin: [M(1,0.36),L(1,0.27),L(0.92,0.15),L(0.68,0.02),L(0.6,0),L(0.4,0),L(0.15,0.11),L(0.08,0.15),L(0,0.27),L(0,0.36),L(0.08,0.56),L(0.5,1),L(0.75,0.76),L(0.92,0.56),L(1,0.36),Z,M(0.3,0.36),L(0.35,0.25),L(0.42,0.21),L(0.58,0.21),L(0.65,0.25),L(0.7,0.36),L(0.65,0.45),L(0.58,0.49),L(0.42,0.49),L(0.35,0.45),L(0.3,0.36),Z],
  gl1: [M(0,0.52),L(0,0.76),L(0.12,1),L(1,1),L(1,0),L(0.12,0),L(0,0.24),L(0,0.52),Z],
  gl2: [M(0,1),L(1,1),L(1,0),L(0.74,0.17),L(0.21,0.72),L(0,1),Z],
  gl3: [M(0.66,0),L(0,0),L(0,1),L(1,1),L(1,0.52),L(0.66,0.52),L(0.66,0),Z],
  gl4: [M(1,0),L(0.37,0),L(0,0.12),L(0,1),L(1,1),L(1,0),Z],
  gl5: [M(0,1),L(0.26,0.79),L(0.79,0.26),L(1,0),L(0,0),L(0,1),Z],
  gl6: [M(1,1),L(1,0),L(0,0),L(0.21,0.26),L(0.74,0.79),L(1,1),Z],
  gl7: [M(0,1),L(0.88,1),L(1,0.5),L(1,0),L(0,0),L(0,1),Z],
  gl8: [M(0,0.92),L(0.24,1),L(0.76,1),L(1,0.92),L(1,0),L(0,0),L(0,0.92),Z],
  plane: [M(1,0.14),L(0.92,0.07),L(0.87,0.07),L(0.65,0.21),L(0.38,0),L(0.25,0.07),L(0.5,0.3),L(0.32,0.42),L(0.13,0.33),L(0,0.42),L(0.33,0.65),L(0.52,0.53),L(0.4,1),L(0.53,0.88),L(0.73,0.42),L(0.97,0.28),L(1,0.14),Z],
  barChart: [M(1,1),L(0,1),L(0,0),L(0.06,0),L(0.06,0.91),L(1,0.91),L(1,1),Z,M(0.32,0.83),L(0.19,0.83),L(0.19,0.5),L(0.32,0.5),L(0.32,0.83),Z,M(0.5,0.83),L(0.38,0.83),L(0.38,0.15),L(0.5,0.15),L(0.5,0.83),Z,M(0.69,0.83),L(0.57,0.83),L(0.57,0.33),L(0.69,0.33),L(0.69,0.83),Z,M(0.88,0.83),L(0.76,0.83),L(0.76,0.07),L(0.88,0.07),L(0.88,0.83),Z],
  priceTag: [M(0.78,0.65),L(0.52,0.98),L(0.47,1),L(0.42,0.98),L(0.05,0.49),B(0.03,0.47,0,0.4,0,0.35),L(0,0.07),B(0,0.04,0.04,0,0.07,0),L(0.29,0),B(0.33,0,0.37,0.02,0.4,0.05),L(0.78,0.53),L(0.78,0.65),Z,M(0.18,0.12),B(0.14,0.12,0.11,0.16,0.11,0.21),B(0.11,0.25,0.14,0.3,0.18,0.3),B(0.21,0.3,0.23,0.25,0.23,0.21),B(0.23,0.16,0.21,0.12,0.18,0.12),Z,M(0.97,0.65),L(0.71,0.98),B(0.71,1,0.68,1,0.67,1),B(0.64,1,0.63,0.98,0.62,0.96),L(0.86,0.65),L(0.86,0.53),L(0.48,0.05),B(0.47,0.02,0.41,0,0.37,0),L(0.49,0),B(0.52,0,0.58,0.02,0.6,0.05),L(0.97,0.53),L(1,0.6),L(0.97,0.65),Z],
  trend: [M(1,1),L(0,1),L(0,0),L(0.06,0),L(0.06,0.91),L(1,0.91),L(1,1),Z,M(0.94,0.38),B(0.94,0.4,0.91,0.41,0.91,0.4),L(0.85,0.31),L(0.54,0.72),B(0.53,0.74,0.53,0.74,0.51,0.72),L(0.4,0.59),L(0.21,0.84),L(0.1,0.72),L(0.38,0.34),B(0.4,0.33,0.41,0.33,0.41,0.34),L(0.53,0.5),L(0.76,0.19),L(0.69,0.1),L(0.71,0.07),L(0.91,0.07),L(0.94,0.1),L(0.94,0.38),Z],
  folder: [M(1,0.83),B(1,0.92,0.94,1,0.86,1),L(0.14,1),B(0.06,1,0,0.92,0,0.83),L(0,0.15),B(0,0.06,0.06,0,0.14,0),L(0.33,0),B(0.41,0,0.47,0.06,0.47,0.15),L(0.47,0.17),L(0.86,0.17),B(0.94,0.17,1,0.25,1,0.34),L(1,0.83),Z],
  search: [M(0.92,1),B(0.89,1,0.88,0.98,0.86,0.97),L(0.66,0.76),B(0.59,0.81,0.52,0.84,0.42,0.84),B(0.19,0.84,0,0.65,0,0.41),B(0,0.19,0.19,0,0.42,0),B(0.66,0,0.84,0.19,0.84,0.41),B(0.84,0.51,0.81,0.59,0.77,0.65),L(0.97,0.86),L(1,0.92),B(1,0.97,0.95,1,0.92,1),Z,M(0.42,0.14),B(0.28,0.14,0.16,0.27,0.16,0.41),B(0.16,0.57,0.28,0.68,0.42,0.68),B(0.58,0.68,0.69,0.57,0.69,0.41),B(0.69,0.27,0.58,0.14,0.42,0.14),Z],
  star: [M(0.98,0.42),L(0.77,0.65),L(0.81,0.97),B(0.81,0.98,0.81,1,0.8,1),L(0.77,1),L(0.5,0.85),L(0.23,1),L(0.2,1),B(0.19,1,0.19,0.98,0.19,0.97),L(0.23,0.65),L(0,0.38),B(0,0.37,0.03,0.37,0.05,0.37),L(0.34,0.32),L(0.47,0.02),L(0.5,0),L(0.53,0.02),L(0.66,0.32),L(0.95,0.37),B(0.97,0.37,1,0.37,1,0.38),L(0.98,0.42),Z],
  phone: [M(0.58,0.58),B(0.5,0.66,0.4,0.74,0.36,0.7),B(0.3,0.64,0.26,0.6,0.14,0.7),B(0,0.8,0.1,0.88,0.16,0.92),B(0.22,1,0.46,0.94,0.7,0.7),B(0.94,0.46,1,0.22,0.94,0.14),B(0.88,0.08,0.82,0,0.72,0.12),B(0.62,0.24,0.66,0.28,0.72,0.34),B(0.76,0.38,0.68,0.48,0.58,0.58)],
  mail: [M(0.04,0.09),L(0.44,0.44),B(0.46,0.47,0.48,0.47,0.5,0.47),B(0.52,0.47,0.54,0.47,0.54,0.44),L(0.94,0.09),B(0.98,0.06,1,0,0.96,0),L(0.04,0),B(0,0,0.02,0.06,0.04,0.09),Z,M(0.96,0.28),B(0.94,0.28,0.56,0.59,0.54,0.62),L(0.44,0.62),B(0.42,0.59,0.06,0.28,0.04,0.28),B(0.02,0.25,0.02,0.28,0.02,0.28),L(0.02,0.93),B(0.02,0.96,0.04,1,0.08,1),L(0.92,1),B(0.96,1,0.98,0.96,0.98,0.93),L(0.98,0.28),B(0.98,0.28,0.98,0.25,0.96,0.28),Z],
  home: [M(0.98,0.52),L(0.54,0.04),B(0.52,0,0.48,0,0.46,0.04),L(0.02,0.52),B(0,0.54,0.02,0.56,0.04,0.56),L(0.13,0.56),L(0.13,0.95),B(0.13,0.98,0.13,1,0.17,1),L(0.38,1),L(0.38,0.61),L(0.61,0.61),L(0.61,1),L(0.83,1),B(0.86,1,0.86,0.98,0.86,0.95),L(0.86,0.56),L(0.96,0.56),B(0.98,0.56,1,0.54,0.98,0.52)],
  globe: [M(0.5,0),B(0.22,0,0,0.22,0,0.5),B(0,0.78,0.22,1,0.5,1),B(0.78,1,1,0.78,1,0.5),B(1,0.22,0.78,0,0.5,0),Z,M(0.92,0.5),B(0.92,0.61,0.89,0.68,0.83,0.76),B(0.81,0.76,0.8,0.72,0.81,0.68),L(0.83,0.54),B(0.83,0.5,0.81,0.43,0.78,0.43),B(0.72,0.43,0.7,0.43,0.67,0.37),B(0.63,0.26,0.78,0.24,0.72,0.18),B(0.7,0.17,0.63,0.24,0.61,0.13),L(0.63,0.11),B(0.8,0.17,0.92,0.31,0.92,0.5),Z,M(0.44,0.09),L(0.39,0.13),B(0.35,0.17,0.33,0.15,0.31,0.18),L(0.24,0.28),L(0.28,0.33),B(0.29,0.31,0.33,0.31,0.37,0.33),B(0.39,0.33,0.57,0.35,0.52,0.5),B(0.5,0.56,0.41,0.54,0.39,0.61),B(0.39,0.63,0.39,0.68,0.37,0.7),B(0.37,0.72,0.39,0.81,0.37,0.81),L(0.28,0.72),L(0.26,0.59),B(0.26,0.56,0.18,0.56,0.18,0.5),B(0.18,0.44,0.22,0.41,0.22,0.39),B(0.2,0.35,0.13,0.35,0.11,0.35),B(0.17,0.2,0.29,0.11,0.44,0.09),Z,M(0.37,0.91),B(0.39,0.89,0.39,0.87,0.43,0.87),L(0.61,0.81),L(0.74,0.85),B(0.67,0.91,0.59,0.92,0.5,0.92),L(0.37,0.91),Z],
};

/* --------------------------------------------------------- text primitives */
const NAV_TXT = { fontFace: BODY, fontSize: 10.5, color: W, valign: 'top' };
const NAV_LEFT = [['Home', 4.340, 0.623], ['Services', 5.665, 0.800], ['About Us', 7.167, 0.863]];
const NAV_RIGHT = [['Home', 5.267, 0.623], ['Services', 6.241, 0.800], ['About Us', 7.392, 0.863]];

/** Body copy: 10.5pt Work Sans, justified, 150% leading (the deck's default). */
function body(s, text, [x, y, w, h], o = {}) {
  s.addText(text, {
    x, y, w, h, fontFace: BODY, fontSize: o.size || 10.5, color: o.color || W,
    align: o.align || 'justify', lineSpacingMultiple: o.solid ? 1 : 1.5,
    bold: !!o.bold, valign: 'top',
  });
}

/** Two-tone display heading, e.g. "Infographic " (green) + "Section" (white). */
function heading(s, green, white, [x, y, w, h], align = 'left', size = 40) {
  s.addText(
    [{ text: green, options: { color: G } }, { text: white, options: { color: W } }],
    { x, y, w, h, fontFace: DISPLAY, fontSize: size, align, valign: 'top' }
  );
}

/** Filled hexagon used as a call-to-action button. */
function hexButton(s, text, [x, y, w, h], fill = G, color = BG, size = 16) {
  s.addText(text, {
    shape: 'hexagon', x, y, w, h, fill: { color: fill }, line: NOLINE,
    fontFace: BODY, fontSize: size, bold: true, color, align: 'center', valign: 'middle',
  });
}

/** Draws one of the ICON outlines scaled into the given box. */
function icon(s, glyph, box, color) {
  freeform(s, box, glyph, { fill: { color } });
}

/** Grey stand-in for a photo; `shape` may be a preset name or 'custGeom'. */
function photo(s, shape, [x, y, w, h], points) {
  if (points) freeform(s, [x, y, w, h], points, { fill: { color: IMG } });
  else s.addShape(shape, { x, y, w, h, fill: { color: IMG }, line: NOLINE });
  s.addText('[image]', {
    x, y: y + h / 2 - 0.16, w, h: 0.32, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 10, color: IMGTX
  });
}

/** The sticky top bar shared by every slide (logo, menu, Log in, Get Started). */
function navBar(s, { logo = G, links = NAV_LEFT, cta = BG } = {}) {
  if (logo) {
    s.addShape('hexagon', { x: 1.161, y: 0.428, w: 0.383, h: 0.330, fill: { color: logo }, line: NOLINE });
    icon(s, ICON.logo, [1.244, 0.506, 0.215, 0.173], BG);
  }
  links.forEach(([t, x, w]) => s.addText(t, { x, y: 0.463, w, h: 0.278, ...NAV_TXT }));
  s.addText('Get Started', {
    shape: 'roundRect', rectRadius: 0, x: 10.696, y: 0.411, w: 1.490, h: 0.330,
    fill: { color: G }, line: NOLINE, fontFace: BODY, fontSize: 10.5, color: cta,
    align: 'center', valign: 'middle',
  });
  s.addText('Log in', { x: 9.902, y: 0.463, w: 0.637, h: 0.278, ...NAV_TXT });
}

/** Centred "Infographic Section" headline used on slides 9-18. */
function infographicTitle(s) {
  heading(s, 'Infographic ', 'Section', [3.250, 1.309, 6.833, 0.774], 'center');
}

/* ------------------------------------------------------------ shared copy */
const T = {
  a: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor',
  b: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut eiusmod tempor incididunt ut consectetur adipiscing elit',
  c: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut',
  d: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt',
  e: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, ',
  f: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut eiusmod tempor',
  g: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut eiusmod tempor incididunt',
  h: 'Lorem ipsum dolor sit amet, consect etuer adipiscing elit. dolor sit etuer',
  i: 'Lorem ipsum dolor ipsum dolor sit',
  j: 'Lorem ipsum dolor sit amet, consectetuer adi piscing elit',
  k: 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttmassa. Fusce posuere, magna sedipsum',
  l: 'Lorem ipsum dolor sit amet, consectetuer adi piscing elit. ',
  m: 'Lorem ipsum dolor sit amet, consectetuer adi elit. ',
  n: 'Lorem ipsum sit dolor sit, ',
  o: 'Lorem ipsum dolor sit amet, adipiscing',
  p: 'Lorem ipsum dolor sit ipsum dolor',
  q: 'Lorem ipsum dolor sit amet, consectetuer adipiscing. elit. Maecenas porttitor.',
  r: 'Lorem dolor sit amet,consectetuer',
  s: 'Lorem ipsum dolor sit amet, adipiscing elit. Maecenas massa. Fusce ipsum dolor sit amet ipsum dolor sit amet',
  intro: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut '
    + 'eiusmod tempor incididunt ut consectetur adipiscing adipiscing elit, sed do eiusmod tempor '
    + 'incididunt ut eiusmod tempor incididunt ut',
};

/* ============================================================ slide builders */

/** 1 — Title slide: giant hex photo behind a full-width green band. */
function slide01(s) {
  photo(s, 'custGeom', [3.033, 0.618, 7.267, 6.265], hexRing(0.1339));
  s.addShape('rect', { x: 0, y: 2.416, w: 13.333, h: 2.669, fill: { color: G }, line: NOLINE });
  navBar(s, { logo: W, links: [] });
  body(s, T.a, [9.721, 3.317, 2.543, 0.867], { color: BG });
  s.addText('GAMER', { x: 0.653, y: 2.995, w: 3.799, h: 1.111, fontFace: DISPLAY, fontSize: 60, color: BG, valign: 'top' });
  s.addText('Esport Presentation Template', { x: 0.653, y: 3.978, w: 3.736, h: 0.370, fontFace: BODY, fontSize: 16, color: BG, valign: 'top' });
  hexButton(s, 'About Us', [1.168, 6.374, 2.144, 0.502]);
  hexButton(s, 'See More', [10.053, 6.374, 2.144, 0.502]);
  photo(s, 'custGeom', [4.279, 1.692, 4.775, 4.116], hexPath(0, 0, 1, 1));
}

/** 2 — Welcome: arrow-shaped photo panel on the left, copy on the right. */
function slide02(s) {
  const wedge = (r) => [M(0, 0), L(r, 0), L(1, 0.5), L(r, 1), L(0, 1), Z];
  freeform(s, [0, 0.017, 6.667, 7.465], wedge(0.6628), { fill: { color: G, transparency: 50 } });
  freeform(s, [0, 0.017, 6.431, 7.465], wedge(0.6505), { fill: { color: G } });
  photo(s, 'custGeom', [0, 0.017, 6.189, 7.465], [M(0, 0), L(0.699, 0), L(1, 0.5), L(0.699, 1), L(0, 1), Z]);
  navBar(s, { logo: null, links: NAV_RIGHT });
  heading(s, 'WELCOME TO ', 'GAMER', [7.372, 1.273, 4.051, 1.447]);
  hexButton(s, 'About Our Planning', [7.457, 3.513, 2.984, 0.502]);
  body(s, T.intro, [7.392, 4.984, 4.903, 1.132]);
}

/** 3 — Vision & Mission: two labelled paragraphs, hex photo on the right. */
function slide03(s) {
  photo(s, 'custGeom', [6.633, 1.379, 5.543, 4.778], [...hexPath(0, 0, 1, 1), ...hexPath(0.176, 0.173, 0.823, 0.819)]);
  navBar(s, { logo: W, cta: INK });
  heading(s, 'Vision & ', 'Mission', [1.084, 1.306, 5.123, 0.774]);
  [['Vision', 2.662, 3.249], ['Mission', 4.698, 5.285]].forEach(([label, hy, ty]) => {
    hexButton(s, label, [1.172, hy, 1.707, 0.383]);
    body(s, T.b, [1.061, ty, 4.445, 0.867]);
  });
  s.addShape('hexagon', { x: 7.331, y: 1.981, w: 4.146, h: 3.574, fill: { color: G, transparency: 50 }, line: NOLINE });
  s.addShape('hexagon', { x: 7.471, y: 2.083, w: 3.867, h: 3.333, fill: { color: G }, line: NOLINE });
  photo(s, 'hexagon', [7.613, 2.206, 3.581, 3.087]);
}

/** 4 — Our Facilities: honeycomb of photos on the left, two blurbs right. */
function slide04(s) {
  photo(s, 'custGeom', [0, 4.528, 2.389, 2.972], [M(0, 0), L(0.610, 0), L(1, 0.626), L(0.768, 1), L(0, 1), Z]);
  photo(s, 'custGeom', [0, 0, 2.514, 4.403], [M(0, 0), L(0.562, 0), L(1, 0.5), L(0.562, 1), L(0, 1), Z]);
  photo(s, 'hexagon', [2.239, 0.582, 1.758, 1.515]);
  photo(s, 'hexagon', [1.559, 2.229, 5.107, 4.403]);
  navBar(s, { logo: null, links: NAV_RIGHT });
  heading(s, 'Our ', 'Facilities', [8.156, 1.280, 4.051, 0.774]);
  [['1. Facilities', 2.807, 3.493], ['2. Facilities', 4.847, 5.533]].forEach(([label, hy, ty]) => {
    hexButton(s, label, [8.258, hy, 2.224, 0.502]);
    body(s, T.c, [8.156, ty, 4.053, 0.601]);
  });
}

/** 5 — Our Best Project: dimmed photo backdrop, green lower half. */
function slide05(s) {
  photo(s, 'rect', [0, 0, 13.333, 7.5]);
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: '000000', transparency: 24 }, line: NOLINE });
  navBar(s, { cta: INK });
  s.addShape('rect', { x: 0, y: 3.75, w: 13.333, h: 3.75, fill: { color: G }, line: NOLINE });
  heading(s, 'Our Best ', 'Project', [3.431, 1.309, 6.470, 0.774], 'center');
  body(s, T.b, [3.603, 2.272, 6.128, 0.601], { align: 'center' });
  [['1. Project Title', 3.274, 3.223], ['2. Project Title', 7.412, 7.366]].forEach(([label, hx, tx]) => {
    hexButton(s, label, [hx, 4.272, 2.369, 0.502], BG, W);
    body(s, T.d, [tx, 5.296, 2.901, 0.867], { color: INK });
  });
  photo(s, 'custGeom', [0, 1.410, 2.795, 4.721], [M(0, 0), L(0.578, 0), L(1, 0.5), L(0.578, 1), L(0, 1), Z]);
  photo(s, 'custGeom', [10.539, 1.390, 2.795, 4.721], [M(0.422, 0), L(1, 0), L(1, 1), L(0.422, 1), L(0, 0.5), Z]);
}

/** 6 — Meet Our Team: two hex portraits with name buttons. */
function slide06(s) {
  navBar(s);
  heading(s, 'Meet Our ', 'Team', [3.431, 1.309, 6.470, 0.774], 'center');
  body(s, T.b, [2.954, 2.222, 7.425, 0.601], { align: 'center' });
  [['Juliana Silva', 1.161, 1.260, 1.388, 4.142], ['Olivia Wilson', 7.065, 7.164, 7.292, 9.934]]
    .forEach(([name, gx, ix, px, tx]) => {
      s.addShape('hexagon', { x: gx, y: 3.636, w: 2.954, h: 2.546, fill: { color: G, transparency: 50 }, line: NOLINE });
      s.addShape('hexagon', { x: ix, y: 3.721, w: 2.754, h: 2.375, fill: { color: G }, line: NOLINE });
      photo(s, 'hexagon', [px, 3.843, 2.506, 2.160]);
      hexButton(s, name, [tx, 4.134, 2.224, 0.502]);
      body(s, T.a, [tx, 5.229, 2.304, 0.867]);
    });
}

/** 7 — Our Best Service: green banner with two arrow-notched ends. */
function slide07(s) {
  const notch = [M(0.098, 0), L(1, 0), L(1, 1), L(0.098, 1), L(0, 0.5), Z];
  freeform(s, [7.288, 3.774, 6.045, 2.375], notch, { fill: { color: G } });
  freeform(s, [0, 3.774, 6.045, 2.375], notch, { fill: { color: G }, flipH: true });
  navBar(s);
  heading(s, 'Our Best ', 'Service', [3.431, 1.309, 6.470, 0.774], 'center');
  body(s, T.b, [3.603, 2.222, 6.128, 0.601], { align: 'center' });
  [['1. Service', ICON.trophy, 0.718, 1.018, 1.392, 0.260, 3.365],
   ['2. Service', ICON.rocket, 10.110, 10.398, 10.783, 0.264, 7.433]]
    .forEach(([label, glyph, hx, gx, tx, gw, px]) => {
      s.addShape('hexagon', { x: hx, y: 4.221, w: 2.224, h: 0.502, fill: { color: BG }, line: NOLINE });
      icon(s, glyph, [gx, 4.335, gw, 0.269], W);
      s.addText(label, { x: tx, y: 4.287, w: 1.390, h: 0.370, fontFace: BODY, fontSize: 16, bold: true, color: W, valign: 'top' });
      body(s, T.e, [hx, 5.010, 2.533, 0.601], { color: INK });
      photo(s, 'hexagon', [px, 3.881, 2.533, 2.152]);
    });
}

/** 8 — Our Portfolio: overlapping hexagonal photo collage. */
function slide08(s) {
  freeform(s, [7.711, 5.526, 2.885, 0.502], [M(0, 0), L(0.991, 0), L(1, 0.1), L(0.922, 1), L(0.087, 1), Z], { fill: { color: G } });
  navBar(s);
  heading(s, 'Our ', 'Portfolio', [1.161, 1.322, 4.051, 0.774]);
  body(s, T.f, [1.072, 2.350, 4.804, 0.601]);
  photo(s, 'hexagon', [9.687, 3.839, 2.539, 2.189]);
  photo(s, 'custGeom', [1.192, 3.800, 6.693, 2.333], [M(0.087, 0), L(0.824, 0), L(1, 1), L(0.087, 1), L(0, 0.5), Z]);
  photo(s, 'custGeom', [6.671, 1.449, 4.658, 4.016],
    [M(0.216, 0), L(0.784, 0), L(1, 0.5), L(0.969, 0.573), L(0.751, 0.573), L(0.625, 0.866), L(0.683, 1), L(0.216, 1), L(0, 0.5), Z]);
}

/** 9 — Pricing cards: three tiers, the middle one raised and highlighted. */
function slide09(s) {
  const bullet = (x, y, text) => {
    s.addShape('ellipse', { x, y: y + 0.095, w: 0.087, h: 0.087, fill: { color: INK }, line: NOLINE });
    s.addText(text, { x: x + 0.106, y, w: 1.883, h: 0.278, fontFace: BODY, fontSize: 10.5, color: INK, valign: 'top' });
  };
  const side = (cardX, title, price) => {
    s.addShape('rect', { x: cardX, y: 2.712, w: 2.750, h: 3.413, fill: { color: G60 }, line: NOLINE });
    s.addText(title, {
      shape: 'roundRect', rectRadius: 0, x: cardX + 0.278, y: 3.047, w: 2.194, h: 0.514,
      fill: { color: G40 }, line: NOLINE, fontFace: BODY, fontSize: 14, bold: true, color: INK, align: 'center', valign: 'middle',
    });
    s.addText([{ text: `$ ${price} `, options: { fontSize: 20 } }, { text: '/Month', options: { fontSize: 11 } }],
      { x: cardX + 0.301, y: 3.653, w: 2.148, h: 0.438, fontFace: BODY, bold: true, color: INK, align: 'center', valign: 'top' });
    bullet(cardX + 0.536, 4.316, 'Your Service Here 01');
    bullet(cardX + 0.536, 4.702, 'Your Service Here 02');
    s.addText('Order Now', {
      shape: 'roundRect', rectRadius: 0, x: cardX + 0.550, y: 5.255, w: 1.649, h: 0.542,
      fill: { color: G40 }, line: NOLINE, fontFace: BODY, fontSize: 11, bold: true, color: INK, align: 'center', valign: 'middle',
    });
  };
  side(2.229, 'Regular Class', 50);
  side(8.354, 'Medium Class', 75);

  s.addShape('rect', { x: 4.979, y: 2.282, w: 3.375, h: 4.233, fill: { color: G }, line: NOLINE, shadow: SHADOW });
  s.addText('Premium Class', {
    shape: 'roundRect', rectRadius: 0, x: 5.353, y: 2.630, w: 2.645, h: 0.607,
    fill: { color: W, transparency: 75 }, line: NOLINE, fontFace: BODY, fontSize: 14, bold: true, color: INK, align: 'center', valign: 'middle',
  });
  s.addText([{ text: '$ 100 ', options: { fontSize: 24 } }, { text: '/Month', options: { fontSize: 12 } }],
    { x: 5.526, y: 3.389, w: 2.148, h: 0.505, fontFace: BODY, bold: true, color: INK, align: 'center', valign: 'top' });
  ['01', '02', '03'].forEach((n, i) => bullet(5.835, 4.103 + i * 0.4185, `Your Service Here ${n}`));
  s.addText('Order Now', {
    shape: 'roundRect', rectRadius: 0, x: 5.548, y: 5.542, w: 2.265, h: 0.654,
    fill: { color: W, transparency: 75 }, line: NOLINE, fontFace: BODY, fontSize: 14, bold: true, color: INK, align: 'center', valign: 'middle',
  });
  navBar(s);
  infographicTitle(s);
}

/** 10 — Six numbered blocks in two columns, alternating green tints. */
function slide10(s) {
  const COLS = [1.255, 7.277];
  const ROWS = [2.664, 3.943, 5.222];
  [['01', 0, 0, G, G], ['03', 0, 1, G60, G40], ['05', 0, 2, G, G],
   ['02', 1, 0, G60, G40], ['04', 1, 1, G, G], ['06', 1, 2, G60, G40]]
    .forEach(([n, c, r, box, ink]) => {
      const bx = COLS[c], by = ROWS[r];
      s.addShape('rect', { x: bx, y: by + 0.071, w: 0.088, h: 0.770, fill: { color: box }, line: NOLINE });
      s.addText(n, {
        shape: 'rect', x: bx + 0.243, y: by, w: 0.950, h: 0.917, fill: { color: box }, line: NOLINE,
        fontFace: BODY, fontSize: 18, bold: true, color: INK, align: 'center', valign: 'middle',
      });
      s.addText('Title Here', { x: bx + 1.433, y: by + 0.013, w: 1.024, h: 0.303, fontFace: BODY, fontSize: 12, bold: true, color: ink, valign: 'top' });
      body(s, T.h, [bx + 1.433, by + 0.286, 3.445, 0.603]);
    });
  navBar(s);
  infographicTitle(s);
}

/** 11 — Interlocking circular-arrow chain with five numbered steps. */
function slide11(s) {
  const RING = [1.161, 3.254, 5.338, 7.436, 9.527];
  RING.forEach((x, i) => {
    const up = i % 2 === 0;
    freeform(s, [x, up ? 2.669 : 3.932, 2.544, 1.279], ICON.arcBg,
      { rotate: up ? 0 : 180, fill: { color: up ? G : G75, transparency: 75 } });
  });
  // painted right-to-left so each bright arrowhead overlaps the pale arc beside it
  [[9.527, 3.602, ICON.arcDn, G], [7.429, 2.671, ICON.arcUp, G60], [5.337, 3.602, ICON.arcDn, G],
   [3.205, 2.669, ICON.arcUp, G60], [1.160, 3.602, ICON.arcDn, G]]
    .forEach(([x, y, glyph, color]) => {
      freeform(s, [x, y, 2.661, y < 3 ? 1.611 : 1.609], glyph, { fill: { color }, shadow: SHADOW });
    });
  ['01', '02', '03', '04', '05'].forEach((n, i) => {
    s.addText(n, {
      x: 1.933 + i * 2.093, y: 3.651 + i * 0.0105, w: 1.0, h: 0.572,
      fontFace: BODY, fontSize: 28, bold: true, color: i % 2 ? G40 : G, align: 'center', valign: 'top',
    });
  });
  for (let i = 0; i < 5; i++) {
    const x = 1.147 + i * 2.310;
    icon(s, ICON.check, [x, 5.631, 0.239, 0.239], G);
    body(s, T.i, [x + 0.339, 5.541, 1.571, 0.603]);
  }
  navBar(s);
  infographicTitle(s);
}

/** 12 — Long right-arrow band pierced by three tilted icon cards. */
function slide12(s) {
  // right arrow with a full-height tail (preset adj1 = 100%)
  freeform(s, [0, 3.446, 12.203, 0.467], [M(0, 0), L(0.981, 0), L(1, 0.5), L(0.981, 1), L(0, 1), Z], { fill: { color: G60 } });
  freeform(s, [0, 3.678, 12.203, 0.235], [M(0, 0), L(1, 0), L(1, 0.007), L(0.981, 1), L(0, 1), Z], { fill: { color: G50 } });
  [[2.648, 2.675, 3.219, ICON.picture, 0.638, 0.508, 3.374, '2.068+', 2.331],
   [5.847, 5.874, 6.457, ICON.people, 0.607, 0.565, 3.358, '3.673+', 5.573],
   [9.047, 9.074, 9.691, ICON.envelope, 0.537, 0.533, 3.374, '1.670+', 8.801]]
    .forEach(([sx, px, gx, glyph, gw, gh, gy, stat, tx], i) => {
      s.addShape('flowChartManualInput', { x: sx, y: 3.107, w: 0.614, h: 1.599, rotate: 13.57, flipH: true, flipV: true, fill: { color: G75 }, line: NOLINE });
      freeform(s, [px, 2.388, 1.612, 1.950], [M(1, 0), L(0, 0.426), L(0, 0.979), L(1, 1), Z],
        { rotate: 346.43, flipH: true, fill: { color: G } });
      icon(s, glyph, [gx, gy, gw, gh], i === 2 ? W : INK);
      s.addText(stat, { x: tx, y: 5.115, w: 1.755, h: 0.505, fontFace: BODY, fontSize: 24, bold: true, color: G, valign: 'top' });
      body(s, T.j, [tx, 5.591, 2.519, 0.627], { size: 11 });
    });
  navBar(s);
  infographicTitle(s);
}

/** 13 — Pinwheel of pentagons with percentages plus a numbered list. */
function slide13(s) {
  s.addShape('pentagon', { x: 2.208, y: 3.411, w: 2.647, h: 2.282, fill: { color: G75 }, line: NOLINE });
  [[3.738, 4.516, 345, G], [2.607, 2.569, 0, G40], [1.445, 4.516, 15, G60]]
    .forEach(([x, y, rot, color]) => s.addShape('pentagon', { x, y, w: 1.848, h: 1.593, rotate: rot, fill: { color }, line: NOLINE }));
  [['53%', 3.016, 3.159], ['86%', 1.852, 5.122], ['42%', 4.140, 5.122]].forEach(([t, x, y]) =>
    s.addText(t, { x, y, w: 1.092, h: 0.505, fontFace: BODY, fontSize: 24, bold: true, color: INK, align: 'center', valign: 'top' }));
  [[3.421, 4.549], [3.370, 4.603], [3.316, 4.654]].forEach(([x, y], i) =>
    icon(s, [ICON.doc1, ICON.doc2, ICON.doc3][i], [x, y, 0.241, 0.302], INK));
  [[2.663, G60], [3.948, G75], [5.234, G]].forEach(([y, color]) => {
    s.addText('01', {
      shape: 'rect', x: 6.335, y, w: 0.702, h: 0.702, fill: { color }, line: NOLINE,
      fontFace: BODY, fontSize: 14, bold: true, color: INK, align: 'center', valign: 'middle',
    });
    body(s, 'Title Here', [7.208, y - 0.113, 3.521, 0.375], { size: 12, bold: true, color });
    body(s, T.k, [7.208, y + 0.328, 5.077, 0.627], { size: 11 });
  });
  navBar(s);
  infographicTitle(s);
}

/** 14 — Hub-and-spoke: central hex avatar with six numbered satellites. */
function slide14(s) {
  [[5.086, 3.874, 270, false], [8.018, 3.874, 270, false]].forEach(([x, y, rot]) =>
    s.addShape('rect', { x, y, w: 0.091, h: 0.975, rotate: rot, fill: { color: G }, line: NOLINE }));
  [[5.457, 3.083, 300], [7.636, 3.083, 60], [5.457, 4.766, 60], [7.636, 4.766, 300]]
    .forEach(([x, y, rot]) => s.addShape('rect', { x, y, w: 0.094, h: 0.882, rotate: rot, fill: { color: G }, line: NOLINE }));
  s.addShape('hexagon', { x: 5.571, y: 3.475, w: 2.058, h: 1.774, fill: { color: G }, line: NOLINE });
  s.addShape('hexagon', { x: 5.677, y: 3.566, w: 1.846, h: 1.591, fill: { color: W }, line: NOLINE, shadow: SHADOW });
  icon(s, ICON.head, [6.365, 3.844, 0.469, 0.615], G);
  icon(s, ICON.torso, [6.160, 4.437, 0.879, 0.442], G);

  // satellites: [label, hexX, hexY, flipH, fill]
  [['01', 4.364, 2.726, false, G60], ['02', 3.778, 3.980, false, G], ['03', 4.397, 5.248, false, G50],
   ['04', 7.917, 2.726, true, G60], ['05', 8.536, 3.980, true, G], ['06', 7.917, 5.248, true, G50]]
    .forEach(([n, x, y, fh, color]) => {
      s.addText(n, {
        shape: 'hexagon', x, y, w: 0.886, h: 0.764, flipH: fh, fill: { color }, line: NOLINE,
        fontFace: BODY, fontSize: 18, bold: true, color: INK, align: 'center', valign: 'middle',
      });
    });
  // captions: [titleColor, y, right-hand?, bodyText, titleX, bodyX, bodyW]
  [[G40, 2.678, T.l, 9.229, 9.229, 3.136], [G, 3.939, T.m, 9.794, 9.794, 2.571], [G50, 5.244, T.l, 9.229, 9.229, 3.136]]
    .forEach(([color, y, text, tx, bx, bw]) => {
      s.addText('Title Here', { x: tx, y, w: 2.133, h: 0.303, fontFace: BODY, fontSize: 12, bold: true, color, valign: 'top' });
      body(s, text, [bx, y + 0.266, bw, 0.627], { size: 11 });
    });
  [[G40, 2.678, T.l, 1.692, 0.690, 3.136], [G, 3.939, T.m, 1.127, 0.690, 2.571], [G50, 5.244, T.l, 1.692, 0.690, 3.136]]
    .forEach(([color, y, text, tx, bx, bw]) => {
      s.addText('Title Here', { x: tx, y, w: 2.133, h: 0.303, fontFace: BODY, fontSize: 12, bold: true, color, align: 'right', valign: 'top' });
      body(s, text, [bx, y + 0.266, bw, 0.627], { size: 11, align: 'right' });
    });
  navBar(s);
  infographicTitle(s);
}

/** 15 — Two "family tree" avatars, each branching into two labelled cards. */
function slide15(s) {
  [[2.090, G75, 2.675, 8.586], [8.001, G50, 8.586, 0]].slice(0, 0); // (kept simple below)
  [{ x: 0, stem: 2.090, tint: G75, card: G, frame: 2.675, ico: 3.185, diagL: 1.491, diagR: 4.112,
     rl: 2.928, rr: 3.508, capL: [1.724, 5.037, '01. Title Here'], capR: [4.345, 4.893, '02. Title Here'] },
   { x: 1, stem: 8.001, tint: G50, card: G75, frame: 8.586, ico: 9.096, diagL: 7.402, diagR: 10.023,
     rl: 8.839, rr: 9.419, capL: [7.635, 5.037, '01. Title Here'], capR: [10.256, 4.893, '02. Title Here'] }]
    .forEach((u) => {
      s.addShape('rect', { x: u.stem, y: 5.893, w: 3.117, h: 0.096, rotate: 90, fill: { color: u.tint }, line: NOLINE });
      s.addShape('rect', { x: u.rl, y: 5.786, w: 0.860, h: 0.096, rotate: 45, fill: { color: u.tint }, line: NOLINE });
      s.addShape('rect', { x: u.rr, y: 5.371, w: 0.860, h: 0.096, rotate: 135, fill: { color: u.tint }, line: NOLINE });
      s.addShape('round2SameRect', { x: u.frame, y: 2.608, w: 1.947, h: 1.872, fill: { color: u.card }, line: NOLINE });
      s.addShape('round2SameRect', { x: u.frame + 0.100, y: 2.704, w: 1.747, h: 1.679, fill: { color: W }, line: NOLINE, shadow: SHADOW });
      icon(s, ICON.head, [u.ico + 0.216, 2.997, 0.495, 0.649], u.card);
      icon(s, ICON.torso, [u.ico, 3.623, 0.927, 0.467], u.card);
      s.addShape('round2DiagRect', { x: u.diagL, y: 4.788, w: 1.819, h: 1.332, flipH: true, fill: { color: u.card }, line: NOLINE });
      s.addShape('round2DiagRect', { x: u.diagR, y: 4.644, w: 1.819, h: 1.332, fill: { color: u.card }, line: NOLINE });
      [u.capL, u.capR].forEach(([cx, cy, label]) => {
        s.addText(label, { x: cx, y: cy, w: 1.353, h: 0.303, fontFace: BODY, fontSize: 12, bold: true, color: W, valign: 'top' });
        body(s, T.n, [cx, cy + 0.231, 1.353, 0.601]);
      });
    });
  navBar(s);
  infographicTitle(s);
}

/** 16 — Zig-zag timeline: four outlined cards linked by bars and triangles. */
function slide16(s) {
  // [barX, barTint, barFlipV, triX, triY, triFlipV, cardY, numY, textY, num, numColor, iconKey]
  [[2.252, G60, false, 2.252, 3.755, false, 2.569, 2.757, 4.569, '01', G40],
   [4.460, G75, true, 4.460, 4.615, true, 5.054, 5.272, 3.413, '02', G75],
   [6.668, G50, false, 6.668, 3.755, false, 2.569, 2.757, 4.569, '03', G50],
   [8.876, G75, true, 8.876, 4.615, true, 5.054, 5.272, 3.413, '04', G75]]
    .forEach(([bx, tint, bflip, tx, ty, tflip, cardY, numY, textY, num, numColor], i) => {
      s.addShape('rect', { x: bx, y: 4.185, w: 2.205, h: 0.335, flipV: bflip, fill: { color: tint }, line: NOLINE, shadow: SOFTSHADOW });
      s.addShape('triangle', { x: tx, y: ty, w: 2.205, h: 0.335, flipV: tflip, fill: { color: tint }, line: NOLINE });
      s.addShape('rect', { x: bx, y: cardY, w: 2.205, h: 1.081, fill: NOLINE, line: { color: tint, width: 0.75 } });
      s.addText(num, { x: bx + 1.043, y: numY, w: 0.998, h: 0.707, fontFace: BODY, fontSize: 36, bold: true, color: numColor, valign: 'top' });
      body(s, T.o, [bx + 0.138, textY, 1.903, 0.627], { size: 11 });
    });
  icon(s, ICON.money, [2.660, 2.934, 0.368, 0.352], G60);
  icon(s, ICON.safeB, [7.093, 2.881, 0.443, 0.457], G50);
  icon(s, ICON.safeA, [7.283, 3.043, 0.120, 0.106], G50);
  icon(s, ICON.pin, [9.406, 5.373, 0.272, 0.379], G75);
  s.addShape('rect', { x: 9.448, y: 5.780, w: 0.175, h: 0.037, fill: { color: G75 }, line: NOLINE });
  // globe-with-plane badge, assembled from nine slivers
  [[ICON.gl1, 4.866, 5.566, 0.111, 0.120], [ICON.gl2, 4.894, 5.432, 0.083, 0.083],
   [ICON.gl3, 5.018, 5.566, 0.134, 0.120], [ICON.gl4, 5.018, 5.404, 0.092, 0.111],
   [ICON.gl5, 5.189, 5.741, 0.097, 0.083], [ICON.gl6, 4.894, 5.741, 0.083, 0.083],
   [ICON.gl7, 5.189, 5.635, 0.125, 0.051], [ICON.gl8, 5.018, 5.741, 0.134, 0.106],
   [ICON.plane, 5.125, 5.418, 0.268, 0.189]]
    .forEach(([glyph, x, y, w, h]) => icon(s, glyph, [x, y, w, h], G75));
  navBar(s);
  infographicTitle(s);
}

/** 17 — Three snipped-corner tiles with progress percentages. */
function slide17(s) {
  [[1.917, G, G40, G, '47%', 1.501, 3.746],
   [5.597, G75, G40, G75, '83%', 5.181, 7.426],
   [9.277, G, G40, G, '19%', 8.861, 11.106]]
    .forEach(([x, outer, wing, mid, pct, lx, rx], i) => {
      const label = [G, G50, G][i];
      s.addShape('snip2DiagRect', { x, y: 2.764 + i * 0.003, w: 2.139, h: 2.139, flipH: true, fill: { color: outer }, line: NOLINE });
      // three chevrons on each side; the middle one is the accent colour
      [[0.317, wing, 1.039], [0.781, mid, 0.726], [1.244, wing, 1.039]].forEach(([dy, color, w], k) => {
        s.addShape('homePlate', { x: rx, y: 2.764 + i * 0.003 + dy, w, h: 0.325, fill: { color }, line: NOLINE });
        s.addShape('homePlate', { x: lx + (k === 1 ? -0.313 : 0), y: 2.770 + i * 0.003 + dy, w: k === 1 ? 1.039 : 0.726, h: 0.325, flipH: true, fill: { color }, line: NOLINE });
      });
      s.addShape('snip2DiagRect', { x: x + 0.153, y: 2.917 + i * 0.003, w: 1.833, h: 1.833, flipH: true, fill: { color: G }, line: NOLINE, shadow: SOFTSHADOW });
      s.addText(pct, { x: x + 0.510, y: 3.229 + i * 0.003, w: 1.185, h: 0.640, fontFace: BODY, fontSize: 32, bold: true, color: INK, valign: 'top' });
      body(s, T.p, [x + 0.306, 3.835 + i * 0.003, 1.527, 0.867], { align: 'center', color: INK });
      s.addText([{ text: `${pct} `, options: { bold: true } }, { text: 'In Progress' }],
        { x: x - 0.058, y: 5.196 + i * 0.003, w: 2.255, h: 0.337, fontFace: BODY, fontSize: 14, color: label, align: 'center', valign: 'top' });
      body(s, T.q, [x - 0.604, 5.516 + i * 0.003, 3.347, 0.603], { align: 'center' });
    });
  navBar(s);
  infographicTitle(s);
}

/** 18 — Descending / ascending chevron of icon diamonds around a centre note. */
function slide18(s) {
  // [x, y, fill, glyph, iconX, iconY, iconW, iconH, arrowX, arrowY, arrowFlipV]
  const NODES = [
    [1.164, 1.400, G, ICON.trend, 1.675, 1.994, 0.656, 0.492, 2.000, 3.079, false],
    [2.720, 2.647, G75, ICON.folder, 3.290, 3.263, 0.539, 0.449, 3.556, 4.298, false],
    [4.276, 3.895, G, ICON.search, 4.847, 4.469, 0.539, 0.531, 5.112, 5.517, false],
    [7.389, 3.895, G, ICON.barChart, 7.901, 4.488, 0.656, 0.492, 8.225, 5.570, false],
    [8.945, 2.647, G75, ICON.star, 9.516, 3.233, 0.539, 0.507, 9.781, 4.325, false],
    [10.502, 1.400, G, ICON.people, 11.033, 1.951, 0.617, 0.578, 11.338, 3.079, false],
  ];
  NODES.forEach(([x, y, fill, glyph, ix, iy, iw, ih, ax, ay]) => {
    s.addShape('diamond', { x, y, w: 1.679, h: 1.679, fill: { color: fill }, line: NOLINE, shadow: { type: 'outer', color: '000000', opacity: 0.33, blur: 8, offset: 3, angle: 90 } });
    icon(s, glyph, [ix, iy, iw, ih], INK);
    s.addShape('line', { x: ax, y: ay, w: 0.007, h: 0.435, line: { color: G, width: 1, endArrowType: 'triangle' } });
  });
  s.addShape('diamond', { x: 5.833, y: 5.142, w: 1.679, h: 1.679, fill: { color: G75 }, line: NOLINE, shadow: { type: 'outer', color: '000000', opacity: 0.33, blur: 8, offset: 3, angle: 90 } });
  icon(s, ICON.priceTag, [6.364, 5.739, 0.617, 0.484], INK);
  s.addShape('line', { x: 6.672, y: 4.692, w: 0, h: 0.450, flipV: true, line: { color: G, width: 1, endArrowType: 'triangle' } });

  // captions: [titleX, titleY, bodyX, colour, flipped?]
  [[1.316, 3.651, 1.135, G], [2.873, 4.844, 2.692, G75], [4.430, 6.037, 4.249, G],
   [10.640, 3.651, 10.459, G], [9.083, 4.844, 8.902, G75], [7.527, 6.037, 7.345, G]]
    .forEach(([tx, ty, bx, color]) => {
      s.addText('Title Here', { x: tx, y: ty, w: 1.375, h: 0.337, fontFace: BODY, fontSize: 14, bold: true, color, align: 'center', valign: 'top' });
      body(s, T.r, [bx, ty + 0.313, 1.737, 0.627], { size: 11, align: 'center' });
    });
  body(s, 'Your Title Here', [5.305, 2.550, 2.724, 0.507], { size: 18, bold: true, color: G75, align: 'center' });
  body(s, T.s, [5.082, 3.096, 3.169, 0.904], { size: 11, align: 'center' });
  navBar(s);
  infographicTitle(s);
}

/** 19 — Contact Us: details over a green lower band, hex photo at right. */
function slide19(s) {
  s.addShape('rect', { x: 0, y: 3.790, w: 13.333, h: 3.710, fill: { color: G }, line: NOLINE });
  navBar(s);
  heading(s, 'Contact', ' Us', [1.073, 1.243, 4.362, 0.774]);
  body(s, T.g, [1.087, 2.524, 4.832, 0.601]);
  [[ICON.phone, 1.167, 4.345, 0.217, 0.215, '+12345678910', 4.314, 1.301],
   [ICON.mail, 1.160, 4.889, 0.225, 0.139, 'yourmail@gmail.com', 4.836, 1.865],
   [ICON.globe, 1.155, 5.357, 0.232, 0.234, 'www.yourwebsite.com', 5.359, 1.953],
   [ICON.home, 1.155, 5.920, 0.225, 0.200, '123 Anywhere ST,, Any City', 5.882, 2.259]]
    .forEach(([glyph, gx, gy, gw, gh, text, ty, tw]) => {
      icon(s, glyph, [gx, gy, gw, gh], BG);
      s.addText(text, { x: 1.599, y: ty, w: tw, h: 0.278, fontFace: BODY, fontSize: 10.5, color: BG, valign: 'top' });
    });
  photo(s, 'custGeom', [6.696, 1.405, 5.470, 4.716], [...hexPath(0, 0, 1, 1), ...hexPath(0.124, 0.130, 0.876, 0.882)]);
  photo(s, 'hexagon', [7.486, 2.104, 3.912, 3.373]);
}

/** 20 — Thank You closer over a dimmed full-bleed photo. */
function slide20(s) {
  photo(s, 'rect', [0, 0, 13.333, 7.5]);
  s.addShape('rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: '000000', transparency: 35 }, line: NOLINE });
  navBar(s);
  heading(s, 'Thank ', 'You', [1.068, 2.138, 5.613, 1.195], 'left', 65);
  s.addText([{ text: 'For', options: { color: W } }, { text: ' watching us', options: { color: G } }],
    { x: 1.042, y: 3.205, w: 5.221, h: 0.808, fontFace: BODY, fontSize: 42, bold: true, valign: 'top' });
  hexButton(s, 'See More Content', [8.893, 5.447, 3.293, 0.643]);
  body(s, T.g, [1.068, 5.541, 4.832, 0.601]);
}

/* --------------------------------------------------------------- assemble */
const BUILDERS = [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
  slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20];

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'GAMER_16x9', width: 13.33333, height: 7.5 });
pptx.layout = 'GAMER_16x9';
pptx.title = 'Gamer — Esport Presentation Template';

BUILDERS.forEach((build) => {
  const slide = pptx.addSlide();
  slide.background = { color: BG };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '10333ad4-3acd-4085-bce1-fd42de315d7b_grok_final.pptx') })
  .then((f) => console.log('wrote', f));
