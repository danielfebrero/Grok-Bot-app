/**
 * Rebecca Florist — presentation template, rebuilt with PptxGenJS.
 *
 *   node 18056061-cca1-443d-91c2-4c391405ec9e_grok_final.js
 *
 * Writes 18056061-cca1-443d-91c2-4c391405ec9e_grok_final.pptx next to this file.
 * Photographic content of the original deck is represented by flat colour
 * placeholders (see phoneMock / monitorMock / worldMap below).
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- palette
const MAUVE = 'AA8483'; // slide background
const PLUM  = '5B4142'; // dark panels
const MIST  = 'DDDEE0'; // light panels
const SAGE  = '9AAE97'; // green half-discs
const MOSS  = '8C9987'; // muted green panels
const WHITE = 'FFFFFF';

const FONT = 'Montserrat Alternates';

// ------------------------------------------------------------- primitives

/** Filled rectangle (no outline); `transparency` is 0-100 percent. */
function rect(s, o) {
  s.addShape('rect', { x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate,
    fill: { color: o.fill, transparency: o.transparency || 0 }, line: { type: 'none' } });
}

/** Filled circle / ellipse. */
function circle(s, o) {
  s.addShape('ellipse', { x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill }, line: { type: 'none' } });
}

/**
 * The deck's signature shape: a half-disc whose flat edge is on the left.
 * Drawn as two cubic beziers (kappa = 0.5523) so it stays a true semicircle
 * at any size; `rotate` turns it to face left / up / down.
 */
function disc(s, o) {
  const k = 0.5523, w = o.w, h = o.h;
  s.addShape('custGeom', {
    x: o.x, y: o.y, w: w, h: h, rotate: o.rotate,
    fill: { color: o.fill }, line: { type: 'none' },
    points: [
      { x: 0, y: 0 },
      { x: w, y: h / 2, curve: { type: 'cubic', x1: k * w, y1: 0, x2: w, y2: (0.5 - k / 2) * h } },
      { x: 0, y: h,     curve: { type: 'cubic', x1: w, y1: (0.5 + k / 2) * h, x2: k * w, y2: h } },
      { close: true },
    ],
  });
}

/** Body / display copy. Every text box in this deck is top-anchored and autofits. */
function txt(s, text, o) {
  s.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h, rotate: o.rotate,
    fontFace: FONT, fontSize: o.fontSize, bold: !!o.bold, italic: !!o.italic,
    color: o.color || '000000', align: o.align || 'left', valign: 'top',
    lineSpacing: o.lineSpacing, paraSpaceBefore: o.paraSpaceBefore,
    wrap: o.wrap !== false, fit: 'resize', margin: [7.2, 7.2, 3.6, 3.6],
  });
}

/** The running header, e.g. "Rebecca Florist | Presentation Template". */
function brand(s, x, y, color) {
  s.addText(
    [
      { text: 'Rebecca Florist ', options: { bold: true, italic: true } },
      { text: '| Presentation Template' },
    ],
    { x: x, y: y, w: 3.092, h: 0.344, fontFace: FONT, fontSize: 10,
      color: color || '000000', valign: 'top', lineSpacing: 20,
      paraSpaceBefore: 12, fit: 'resize', margin: [7.2, 7.2, 3.6, 3.6] }
  );
}

/** Small bold label above a 9pt supporting sentence — repeated all over the deck. */
function note(s, o) {
  txt(s, o.label, { x: o.x, y: o.y, w: o.labelW || 1.706, h: 0.353,
    fontSize: o.labelSize || 11, bold: true, color: o.color, align: o.align,
    lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, o.body, { x: o.x, y: o.y + (o.gap || 0.332), w: o.w, h: o.h,
    fontSize: 9, color: o.color, align: o.align, lineSpacing: 20, paraSpaceBefore: 12 });
}

// --------------------------------------------------------- image stand-ins

const CAPTION = '3E6B66'; // caption colour used inside the image placeholders

/**
 * Placeholder for the phone photo mock-ups on slides 26 and 27: a dark rounded
 * body with a flat mint "screenshot" and an optional front-camera notch.
 */
function phoneMock(s, o) {
  const b = o.bezel;
  s.addShape('roundRect', { x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: o.radius,
    fill: { color: '111111' }, line: { type: 'none' } });
  s.addShape('roundRect', { x: o.x + b, y: o.y + b, w: o.w - 2 * b, h: o.h - 2 * b,
    rectRadius: o.radius * 0.6, fill: { color: o.screen }, line: { type: 'none' } });
  if (o.notch) {
    s.addShape('roundRect', { x: o.x + o.w / 2 - 0.4, y: o.y + b - 0.01, w: 0.8, h: 0.13,
      rectRadius: 0.07, fill: { color: '111111' }, line: { type: 'none' } });
  }
  txt(s, '[image]', { x: o.x, y: o.y + o.h / 2 - 0.14, w: o.w, h: 0.28,
    fontSize: 10, color: CAPTION, align: 'center' });
}

/** Placeholder for the desktop-monitor photo mock-up on slide 28. */
function monitorMock(s, o) {
  rect(s, { x: o.x, y: o.y, w: o.w, h: o.h, fill: '111111' });                     // bezel
  rect(s, { x: o.x + 0.25, y: o.y + 0.28, w: o.w - 0.5, h: o.h - 0.56, fill: o.screen });
  rect(s, { x: o.x, y: o.y + o.h, w: o.w, h: 0.2, fill: 'C0C0C0' });               // chin
  rect(s, { x: o.x + o.w / 2 - 0.7, y: o.y + o.h + 0.2, w: 1.4, h: 0.83, fill: 'AEAEAE' }); // neck
  rect(s, { x: o.x + o.w / 2 - 0.72, y: o.y + o.h + 1.03, w: 1.44, h: 0.21, fill: 'E4E4E4' }); // foot
  txt(s, '[image]', { x: o.x, y: o.y + o.h / 2 - 0.16, w: o.w, h: 0.32,
    fontSize: 12, color: CAPTION, align: 'center' });
}

/** Outlined social-media glyph stand-in used on the contact slide. */
function iconDot(s, x, y, glyph) {
  s.addShape('ellipse', { x: x, y: y, w: 0.346, h: 0.346,
    fill: { type: 'none' }, line: { color: '000000', width: 1.75 } });
  s.addText(glyph, { x: x, y: y, w: 0.346, h: 0.346, fontFace: FONT, fontSize: 11,
    bold: true, color: '000000', align: 'center', valign: 'middle', margin: 0 });
}

/**
 * Slide 29's world map, traced from the original artwork. Each entry is a flat
 * [x0,y0, x1,y1, ...] polygon in 0..1 units of the map's bounding box: WORLD_LAND
 * is the white land mass, WORLD_HIGHLIGHT the darker call-out countries on top.
 */
const WORLD_LAND = [
  [0.999, 0.33, 0.985, 0.321, 0.959, 0.287, 0.944, 0.285, 0.943, 0.298, 0.937, 0.291, 0.916,
   0.291, 0.912, 0.274, 0.896, 0.274, 0.886, 0.256, 0.864, 0.247, 0.858, 0.252, 0.858, 0.265,
   0.838, 0.263, 0.837, 0.27, 0.833, 0.272, 0.828, 0.243, 0.824, 0.237, 0.814, 0.234, 0.813,
   0.245, 0.806, 0.245, 0.801, 0.235, 0.787, 0.235, 0.778, 0.229, 0.777, 0.223, 0.788, 0.207,
   0.788, 0.198, 0.78, 0.184, 0.77, 0.184, 0.763, 0.166, 0.755, 0.172, 0.75, 0.189, 0.731, 0.195,
   0.714, 0.212, 0.711, 0.232, 0.696, 0.234, 0.695, 0.255, 0.683, 0.257, 0.68, 0.247, 0.68,
   0.257, 0.675, 0.26, 0.673, 0.247, 0.664, 0.246, 0.658, 0.27, 0.661, 0.303, 0.64, 0.287, 0.64,
   0.301, 0.62, 0.3, 0.607, 0.314, 0.604, 0.323, 0.599, 0.321, 0.601, 0.307, 0.594, 0.303, 0.596,
   0.326, 0.587, 0.325, 0.585, 0.309, 0.56, 0.287, 0.558, 0.278, 0.552, 0.272, 0.542, 0.273,
   0.538, 0.283, 0.526, 0.29, 0.504, 0.345, 0.488, 0.371, 0.491, 0.402, 0.497, 0.402, 0.5, 0.396,
   0.506, 0.401, 0.505, 0.408, 0.497, 0.413, 0.497, 0.437, 0.488, 0.441, 0.484, 0.454, 0.478,
   0.455, 0.478, 0.445, 0.466, 0.42, 0.468, 0.408, 0.464, 0.406, 0.465, 0.4, 0.458, 0.405, 0.459,
   0.427, 0.447, 0.436, 0.447, 0.45, 0.451, 0.45, 0.456, 0.444, 0.458, 0.428, 0.465, 0.431,
   0.466, 0.439, 0.461, 0.44, 0.458, 0.462, 0.469, 0.459, 0.469, 0.47, 0.461, 0.472, 0.471,
   0.488, 0.471, 0.499, 0.468, 0.505, 0.449, 0.506, 0.45, 0.539, 0.457, 0.542, 0.457, 0.551,
   0.428, 0.614, 0.428, 0.656, 0.439, 0.679, 0.452, 0.692, 0.487, 0.684, 0.491, 0.693, 0.501,
   0.696, 0.499, 0.717, 0.507, 0.732, 0.512, 0.755, 0.507, 0.792, 0.514, 0.81, 0.516, 0.833,
   0.524, 0.853, 0.525, 0.867, 0.528, 0.871, 0.552, 0.86, 0.564, 0.838, 0.564, 0.824, 0.572,
   0.816, 0.57, 0.797, 0.586, 0.776, 0.582, 0.731, 0.606, 0.691, 0.614, 0.665, 0.613, 0.66,
   0.595, 0.666, 0.593, 0.657, 0.633, 0.627, 0.639, 0.612, 0.63, 0.603, 0.631, 0.596, 0.657,
   0.598, 0.669, 0.618, 0.674, 0.618, 0.685, 0.671, 0.693, 0.671, 0.696, 0.684, 0.699, 0.679,
   0.694, 0.669, 0.696, 0.642, 0.715, 0.614, 0.727, 0.612, 0.735, 0.633, 0.735, 0.642, 0.744,
   0.64, 0.745, 0.677, 0.751, 0.684, 0.752, 0.696, 0.738, 0.691, 0.768, 0.743, 0.791, 0.748,
   0.766, 0.736, 0.766, 0.724, 0.76, 0.714, 0.759, 0.691, 0.75, 0.678, 0.748, 0.664, 0.75, 0.656,
   0.757, 0.657, 0.764, 0.674, 0.775, 0.66, 0.774, 0.644, 0.767, 0.633, 0.766, 0.624, 0.777,
   0.616, 0.774, 0.63, 0.777, 0.63, 0.78, 0.616, 0.797, 0.607, 0.81, 0.581, 0.804, 0.547, 0.812,
   0.539, 0.802, 0.538, 0.799, 0.528, 0.809, 0.521, 0.809, 0.528, 0.819, 0.524, 0.823, 0.551,
   0.829, 0.549, 0.831, 0.539, 0.826, 0.523, 0.833, 0.51, 0.849, 0.498, 0.86, 0.473, 0.861,
   0.454, 0.866, 0.451, 0.865, 0.505, 0.861, 0.507, 0.86, 0.532, 0.832, 0.561, 0.835, 0.569,
   0.838, 0.561, 0.861, 0.55, 0.865, 0.529, 0.864, 0.511, 0.875, 0.505, 0.866, 0.493, 0.869,
   0.485, 0.867, 0.472, 0.872, 0.466, 0.868, 0.435, 0.866, 0.44, 0.853, 0.436, 0.85, 0.424,
   0.866, 0.395, 0.901, 0.395, 0.9, 0.387, 0.905, 0.376, 0.913, 0.373, 0.918, 0.388, 0.907,
   0.406, 0.902, 0.424, 0.907, 0.455, 0.923, 0.418, 0.921, 0.401, 0.925, 0.389, 0.937, 0.384,
   0.944, 0.388, 0.953, 0.373, 0.969, 0.366, 0.966, 0.347, 0.972, 0.336, 0.98, 0.336, 0.99,
   0.348, 0.991, 0.338],
  [0.028, 0.281, 0.021, 0.298, 0.014, 0.304, 0.026, 0.329, 0.02, 0.325, 0.011, 0.335, 0.015,
   0.344, 0.03, 0.348, 0.029, 0.356, 0.02, 0.361, 0.016, 0.375, 0.018, 0.384, 0.027, 0.389,
   0.028, 0.398, 0.039, 0.398, 0.039, 0.408, 0.02, 0.431, 0.034, 0.422, 0.043, 0.408, 0.05,
   0.413, 0.053, 0.408, 0.05, 0.404, 0.05, 0.391, 0.055, 0.388, 0.057, 0.393, 0.065, 0.387,
   0.066, 0.38, 0.104, 0.402, 0.121, 0.448, 0.12, 0.462, 0.131, 0.472, 0.134, 0.532, 0.142,
   0.552, 0.151, 0.561, 0.157, 0.587, 0.171, 0.609, 0.165, 0.592, 0.169, 0.586, 0.188, 0.63,
   0.232, 0.653, 0.239, 0.669, 0.251, 0.679, 0.258, 0.674, 0.261, 0.697, 0.251, 0.717, 0.25,
   0.737, 0.264, 0.772, 0.281, 0.794, 0.277, 0.863, 0.272, 0.884, 0.269, 0.916, 0.272, 0.922,
   0.266, 0.937, 0.267, 0.973, 0.282, 0.999, 0.293, 0.995, 0.284, 0.981, 0.283, 0.962, 0.292,
   0.944, 0.288, 0.931, 0.298, 0.913, 0.302, 0.891, 0.315, 0.887, 0.315, 0.871, 0.324, 0.869,
   0.334, 0.85, 0.342, 0.82, 0.361, 0.807, 0.366, 0.788, 0.366, 0.771, 0.377, 0.75, 0.377, 0.739,
   0.364, 0.724, 0.336, 0.713, 0.329, 0.69, 0.316, 0.687, 0.301, 0.666, 0.286, 0.666, 0.276,
   0.658, 0.267, 0.664, 0.264, 0.673, 0.247, 0.673, 0.242, 0.664, 0.244, 0.646, 0.23, 0.642,
   0.234, 0.617, 0.226, 0.618, 0.222, 0.63, 0.209, 0.63, 0.204, 0.614, 0.205, 0.587, 0.212,
   0.577, 0.225, 0.58, 0.228, 0.573, 0.244, 0.577, 0.25, 0.599, 0.253, 0.592, 0.261, 0.592,
   0.253, 0.591, 0.25, 0.567, 0.264, 0.549, 0.269, 0.52, 0.28, 0.514, 0.281, 0.501, 0.291, 0.495,
   0.294, 0.502, 0.308, 0.489, 0.307, 0.484, 0.299, 0.486, 0.294, 0.481, 0.296, 0.468, 0.303,
   0.468, 0.303, 0.461, 0.315, 0.455, 0.311, 0.477, 0.328, 0.483, 0.326, 0.468, 0.319, 0.462,
   0.32, 0.44, 0.304, 0.418, 0.296, 0.386, 0.291, 0.398, 0.286, 0.401, 0.282, 0.395, 0.281,
   0.379, 0.271, 0.367, 0.259, 0.366, 0.258, 0.397, 0.263, 0.418, 0.255, 0.429, 0.255, 0.453,
   0.248, 0.446, 0.246, 0.427, 0.219, 0.411, 0.212, 0.396, 0.215, 0.371, 0.226, 0.349, 0.235,
   0.344, 0.234, 0.354, 0.238, 0.36, 0.242, 0.353, 0.252, 0.354, 0.238, 0.334, 0.239, 0.327,
   0.249, 0.318, 0.248, 0.286, 0.262, 0.29, 0.265, 0.3, 0.262, 0.317, 0.266, 0.314, 0.267, 0.305,
   0.273, 0.31, 0.269, 0.336, 0.259, 0.339, 0.258, 0.345, 0.27, 0.347, 0.284, 0.366, 0.291, 0.37,
   0.291, 0.364, 0.295, 0.361, 0.286, 0.329, 0.291, 0.327, 0.298, 0.339, 0.303, 0.322, 0.286,
   0.303, 0.288, 0.292, 0.285, 0.279, 0.274, 0.265, 0.269, 0.266, 0.258, 0.234, 0.237, 0.239,
   0.236, 0.233, 0.228, 0.241, 0.228, 0.276, 0.239, 0.285, 0.236, 0.309, 0.231, 0.314, 0.231,
   0.301, 0.22, 0.288, 0.214, 0.261, 0.216, 0.247, 0.223, 0.233, 0.219, 0.228, 0.211, 0.23,
   0.209, 0.248, 0.204, 0.246, 0.205, 0.234, 0.198, 0.233, 0.192, 0.248, 0.2, 0.266, 0.209,
   0.265, 0.208, 0.285, 0.203, 0.283, 0.199, 0.291, 0.204, 0.298, 0.201, 0.312, 0.183, 0.304,
   0.184, 0.298, 0.191, 0.299, 0.195, 0.285, 0.185, 0.27, 0.185, 0.237, 0.166, 0.25, 0.159,
   0.247, 0.15, 0.226, 0.139, 0.223, 0.13, 0.225, 0.131, 0.238, 0.127, 0.261, 0.135, 0.272,
   0.145, 0.259, 0.148, 0.277, 0.153, 0.278, 0.15, 0.286, 0.152, 0.298, 0.138, 0.287, 0.126,
   0.288, 0.121, 0.279, 0.12, 0.285, 0.094, 0.298, 0.077, 0.283, 0.055, 0.279, 0.042, 0.268],
  [0.385, 0, 0.372, 0, 0.364, 0.018, 0.354, 0.017, 0.345, 0.036, 0.344, 0.056, 0.335, 0.044,
   0.33, 0.057, 0.318, 0.052, 0.303, 0.062, 0.302, 0.078, 0.287, 0.107, 0.293, 0.118, 0.293,
   0.125, 0.272, 0.149, 0.273, 0.158, 0.283, 0.167, 0.277, 0.177, 0.285, 0.194, 0.301, 0.191,
   0.312, 0.203, 0.323, 0.247, 0.32, 0.263, 0.327, 0.268, 0.323, 0.29, 0.334, 0.291, 0.326,
   0.316, 0.326, 0.33, 0.338, 0.374, 0.354, 0.386, 0.355, 0.362, 0.364, 0.335, 0.376, 0.327,
   0.386, 0.307, 0.401, 0.299, 0.413, 0.276, 0.41, 0.239, 0.416, 0.237, 0.416, 0.225, 0.42,
   0.224, 0.415, 0.182, 0.422, 0.177, 0.419, 0.138, 0.424, 0.105, 0.439, 0.079, 0.438, 0.066,
   0.43, 0.061, 0.412, 0.075, 0.41, 0.056, 0.407, 0.062, 0.398, 0.056, 0.398, 0.048, 0.411,
   0.045, 0.415, 0.034, 0.399, 0.006],
  [0.298, 0.028, 0.274, 0.017, 0.255, 0.021, 0.247, 0.03, 0.245, 0.04, 0.238, 0.038, 0.233, 0.05,
   0.221, 0.061, 0.228, 0.097, 0.223, 0.096, 0.219, 0.08, 0.213, 0.082, 0.207, 0.11, 0.212,
   0.127, 0.215, 0.128, 0.218, 0.15, 0.234, 0.153, 0.228, 0.186, 0.252, 0.191, 0.258, 0.181,
   0.254, 0.172, 0.265, 0.147, 0.265, 0.128, 0.277, 0.116, 0.281, 0.098, 0.292, 0.075, 0.288,
   0.067, 0.303, 0.045],
  [0.896, 0.827, 0.885, 0.801, 0.877, 0.794, 0.867, 0.758, 0.863, 0.788, 0.848, 0.777, 0.851,
   0.763, 0.839, 0.761, 0.831, 0.776, 0.82, 0.775, 0.809, 0.797, 0.791, 0.807, 0.788, 0.816,
   0.793, 0.854, 0.792, 0.867, 0.796, 0.871, 0.836, 0.853, 0.844, 0.858, 0.848, 0.869, 0.858,
   0.875, 0.863, 0.887, 0.877, 0.893, 0.887, 0.885, 0.895, 0.858],
  [0.538, 0.097, 0.523, 0.103, 0.515, 0.122, 0.512, 0.122, 0.51, 0.114, 0.503, 0.122, 0.505,
   0.14, 0.513, 0.16, 0.512, 0.169, 0.518, 0.181, 0.526, 0.145, 0.533, 0.137, 0.528, 0.124,
   0.544, 0.124, 0.549, 0.113],
  [0.734, 0.085, 0.726, 0.103, 0.73, 0.109, 0.731, 0.125, 0.743, 0.141, 0.751, 0.141, 0.749,
   0.158, 0.763, 0.151, 0.764, 0.142, 0.758, 0.131, 0.75, 0.132, 0.75, 0.118, 0.739, 0.08],
  [0.662, 0.18, 0.643, 0.191, 0.628, 0.212, 0.617, 0.259, 0.623, 0.274, 0.632, 0.276, 0.626,
   0.263, 0.626, 0.252, 0.634, 0.225, 0.645, 0.207, 0.661, 0.193],
  [0.207, 0.177, 0.218, 0.199, 0.218, 0.21, 0.212, 0.203, 0.208, 0.207, 0.209, 0.217, 0.218,
   0.213, 0.228, 0.221, 0.247, 0.221, 0.253, 0.215, 0.249, 0.201, 0.229, 0.204, 0.223, 0.194,
   0.221, 0.182],
  [0.835, 0.715, 0.841, 0.727, 0.854, 0.732, 0.855, 0.746, 0.869, 0.75, 0.871, 0.744, 0.876,
   0.744, 0.889, 0.757, 0.873, 0.728, 0.857, 0.719, 0.846, 0.724, 0.844, 0.715],
  [0.803, 0.688, 0.796, 0.682, 0.792, 0.693, 0.775, 0.705, 0.779, 0.723, 0.794, 0.727, 0.801,
   0.706, 0.799, 0.693],
  [0.182, 0.198, 0.176, 0.197, 0.172, 0.182, 0.172, 0.203, 0.168, 0.204, 0.155, 0.188, 0.15,
   0.21, 0.161, 0.21, 0.161, 0.221, 0.18, 0.212],
  [0.611, 0.765, 0.596, 0.787, 0.594, 0.809, 0.599, 0.823, 0.605, 0.819, 0.613, 0.779],
  [0.436, 0.34, 0.434, 0.326, 0.419, 0.332, 0.412, 0.327, 0.407, 0.334, 0.413, 0.352, 0.42, 0.354],
  [0.966, 0.886, 0.961, 0.887, 0.951, 0.869, 0.956, 0.882, 0.954, 0.895, 0.957, 0.904, 0.95,
   0.903, 0.934, 0.934, 0.943, 0.935],
  [0.854, 0.197, 0.853, 0.208, 0.858, 0.219, 0.872, 0.215, 0.873, 0.202, 0.861, 0.194],
  [0.153, 0.166, 0.145, 0.168, 0.135, 0.193, 0.139, 0.197, 0.151, 0.185],
  [0.202, 0.184, 0.191, 0.191, 0.191, 0.202, 0.196, 0.202, 0.198, 0.213, 0.203, 0.211],
];
const WORLD_HIGHLIGHT = [
  [0.085, 0.29, 0.085, 0.384, 0.095, 0.396, 0.101, 0.391, 0.116, 0.42, 0.115, 0.433, 0.121,
   0.448, 0.12, 0.462, 0.128, 0.471, 0.212, 0.468, 0.231, 0.473, 0.247, 0.492, 0.245, 0.512,
   0.267, 0.493, 0.277, 0.493, 0.282, 0.48, 0.287, 0.48, 0.288, 0.49, 0.294, 0.502, 0.308, 0.489,
   0.307, 0.484, 0.296, 0.484, 0.296, 0.468, 0.303, 0.468, 0.303, 0.461, 0.312, 0.455, 0.315,
   0.462, 0.311, 0.477, 0.328, 0.483, 0.326, 0.468, 0.319, 0.462, 0.32, 0.44, 0.304, 0.418,
   0.296, 0.386, 0.291, 0.398, 0.286, 0.401, 0.282, 0.395, 0.281, 0.379, 0.271, 0.367, 0.259,
   0.366, 0.258, 0.397, 0.263, 0.418, 0.255, 0.429, 0.255, 0.453, 0.248, 0.446, 0.246, 0.427,
   0.219, 0.411, 0.212, 0.396, 0.215, 0.371, 0.226, 0.349, 0.235, 0.344, 0.234, 0.354, 0.238,
   0.36, 0.242, 0.353, 0.252, 0.354, 0.238, 0.334, 0.239, 0.327, 0.249, 0.318, 0.248, 0.286,
   0.262, 0.29, 0.265, 0.3, 0.262, 0.317, 0.266, 0.314, 0.267, 0.305, 0.273, 0.31, 0.269, 0.336,
   0.259, 0.339, 0.259, 0.348, 0.27, 0.347, 0.278, 0.361, 0.291, 0.37, 0.291, 0.364, 0.295,
   0.361, 0.286, 0.329, 0.291, 0.327, 0.298, 0.339, 0.303, 0.322, 0.286, 0.303, 0.288, 0.292,
   0.285, 0.279, 0.274, 0.265, 0.269, 0.266, 0.258, 0.234, 0.237, 0.239, 0.236, 0.233, 0.228,
   0.241, 0.228, 0.276, 0.239, 0.285, 0.234, 0.314, 0.231, 0.313, 0.229, 0.298, 0.22, 0.288,
   0.214, 0.261, 0.216, 0.247, 0.223, 0.233, 0.219, 0.228, 0.211, 0.23, 0.209, 0.248, 0.204,
   0.246, 0.205, 0.234, 0.198, 0.233, 0.192, 0.248, 0.2, 0.266, 0.209, 0.265, 0.208, 0.285,
   0.203, 0.283, 0.199, 0.291, 0.204, 0.298, 0.201, 0.312, 0.183, 0.304, 0.184, 0.298, 0.191,
   0.299, 0.195, 0.285, 0.185, 0.27, 0.185, 0.237, 0.166, 0.25, 0.159, 0.247, 0.15, 0.226, 0.139,
   0.223, 0.13, 0.225, 0.131, 0.238, 0.127, 0.261, 0.135, 0.272, 0.145, 0.259, 0.148, 0.277,
   0.153, 0.278, 0.15, 0.286, 0.152, 0.298, 0.138, 0.287, 0.126, 0.288, 0.121, 0.279, 0.12,
   0.285, 0.102, 0.29, 0.098, 0.298],
  [0.845, 0.475, 0.835, 0.479, 0.824, 0.463, 0.82, 0.444, 0.808, 0.441, 0.804, 0.462, 0.796,
   0.466, 0.794, 0.473, 0.801, 0.475, 0.803, 0.485, 0.783, 0.495, 0.779, 0.51, 0.759, 0.515,
   0.74, 0.51, 0.736, 0.501, 0.726, 0.495, 0.724, 0.483, 0.715, 0.47, 0.696, 0.497, 0.695, 0.515,
   0.678, 0.529, 0.693, 0.552, 0.692, 0.567, 0.72, 0.586, 0.741, 0.578, 0.747, 0.589, 0.744,
   0.604, 0.754, 0.617, 0.757, 0.609, 0.767, 0.608, 0.777, 0.617, 0.776, 0.631, 0.78, 0.616,
   0.797, 0.607, 0.81, 0.581, 0.804, 0.547, 0.812, 0.539, 0.802, 0.538, 0.799, 0.528, 0.805,
   0.521, 0.81, 0.528, 0.834, 0.506, 0.835, 0.494, 0.841, 0.49],
  [0.298, 0.028, 0.274, 0.017, 0.255, 0.021, 0.247, 0.03, 0.245, 0.04, 0.238, 0.038, 0.233, 0.05,
   0.221, 0.061, 0.228, 0.097, 0.223, 0.096, 0.219, 0.08, 0.213, 0.082, 0.207, 0.11, 0.212,
   0.127, 0.215, 0.128, 0.218, 0.15, 0.234, 0.153, 0.228, 0.186, 0.252, 0.191, 0.258, 0.181,
   0.254, 0.172, 0.265, 0.147, 0.265, 0.128, 0.277, 0.116, 0.281, 0.098, 0.292, 0.075, 0.288,
   0.067, 0.303, 0.045],
  [0.896, 0.827, 0.885, 0.801, 0.877, 0.794, 0.867, 0.758, 0.863, 0.788, 0.848, 0.777, 0.851,
   0.763, 0.839, 0.761, 0.831, 0.776, 0.82, 0.775, 0.809, 0.797, 0.791, 0.807, 0.788, 0.816,
   0.793, 0.854, 0.792, 0.867, 0.796, 0.871, 0.836, 0.853, 0.844, 0.858, 0.848, 0.869, 0.858,
   0.875, 0.863, 0.887, 0.877, 0.893, 0.887, 0.885, 0.895, 0.858],
  [0.324, 0.825, 0.321, 0.833, 0.314, 0.834, 0.314, 0.823, 0.299, 0.809, 0.291, 0.809, 0.285,
   0.82, 0.276, 0.903, 0.277, 0.937, 0.272, 0.955, 0.277, 0.973, 0.284, 0.973, 0.283, 0.962,
   0.292, 0.944, 0.288, 0.931, 0.298, 0.913, 0.301, 0.893, 0.316, 0.884, 0.312, 0.869, 0.314,
   0.847, 0.325, 0.831],
  [0.506, 0.561, 0.501, 0.574, 0.502, 0.602, 0.512, 0.611, 0.52, 0.608, 0.539, 0.625, 0.542,
   0.622, 0.542, 0.567, 0.531, 0.563, 0.525, 0.574],
  [0.207, 0.177, 0.218, 0.199, 0.218, 0.21, 0.212, 0.203, 0.208, 0.207, 0.209, 0.217, 0.218,
   0.213, 0.228, 0.221, 0.247, 0.221, 0.253, 0.215, 0.249, 0.201, 0.229, 0.204, 0.223, 0.194,
   0.221, 0.182],
  [0.182, 0.198, 0.176, 0.197, 0.172, 0.182, 0.172, 0.203, 0.168, 0.204, 0.155, 0.188, 0.15,
   0.21, 0.161, 0.21, 0.161, 0.221, 0.18, 0.212],
  [0.153, 0.166, 0.145, 0.168, 0.135, 0.193, 0.139, 0.197, 0.151, 0.185],
  [0.202, 0.184, 0.191, 0.191, 0.191, 0.202, 0.196, 0.202, 0.198, 0.213, 0.203, 0.211],
];

function worldMap(s, o) {
  const draw = (poly, color) => {
    const points = [];
    for (let i = 0; i < poly.length; i += 2) points.push({ x: o.w * poly[i], y: o.h * poly[i + 1] });
    points.push({ close: true });
    s.addShape('custGeom', { x: o.x, y: o.y, w: o.w, h: o.h,
      fill: { color: color }, line: { type: 'none' }, points: points });
  };
  WORLD_LAND.forEach((p) => draw(p, WHITE));
  WORLD_HIGHLIGHT.forEach((p) => draw(p, PLUM));
}
// ------------------------------------------------- shared body copy
const T1 = 'Ut wisi enim ad minim veniam, quis nostrud exerci tation iriure dolor in hendrerit nam';
const T2 = 'Ut wisi enim ad minim veniam, quis hendrerit nam odio moliestie mirum';
const T3 =
  'Duis autem vel eum iriure dolor in hendrerit in velit esse molestie consequat, ' +
  'vel illum dolore eu feugiat';
const T4 = 'PLACEHOLDER';
const T5 =
  'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum ' +
  'iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel illum ' +
  'dolore eu feugiat';
const T6 = 'Ut wisi enim ad minim veniam, quis nostrud tation iriure';
const T7 =
  'Ut wisi enim ad minim veniam, quis nostrud exerci tation iriure dolor in ' +
  'hendrerit nam odio';
const T8 = 'Mirum est notare quam littera gothica, quam nunc putamus parum claram, anteposuerit';
const T9 = 'Duis autem vel eum iriure dolor in hendrerit in vulputate velit';
const T10 = 'Mirum est notare quam littera gothica, quam nunc putamus';
const T11 =
  'What each of us believes in is up to us, but life is impossible without ' +
  'believing in something';
const T12 = 'Mirum est notare quam littera gothica, quam nunc putamus parum claram,';
const T13 = 'You will face many defeats in life, but don’t let yourself be defeated';
const T14 =
  'Duis autem vel eum iriure dolor in hendrerit in velit esse molestie consequat, ' +
  'vel illum';
const T15 =
  'Ut wisi enim ad minim veniam, quis nostrud exerci tation duis autem vel eum ' +
  'iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel illum ' +
  'dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio. Duis ' +
  'PLACEHOLDER' +
  'consequat, vel';
const T16 =
  'Ut wisi enim ad minim veniam, quis nostrud exerci tation iriure dolor in ' +
  'hendrerit nam odio moliestie mirum lorem ipsum volutpat sit';
const T17 =
  'Duis autem vel eum iriure dolor in hendrerit in velit esse molestie consequat, ' +
  'PLACEHOLDER';
const T18 =
  'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum ' +
  'iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel illum ' +
  'dolore eu feugiat nulla facilisis at vero eros et accumsan et iusto odio.';
const T19 = 'PLACEHOLDER';
const T20 =
  'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum ' +
  'iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel illum ' +
  'dolore eu feugiat nulla facilisis at vero eros et accumsan odio.';
const T21 =
  'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum ' +
  'iriure dolor in hendrerit in';

// ----------------------------------------------------------- slides

// Slide 1 — Rebecca cover
function slide1(s) {
  disc(s, { x: 2, y: 2.052, w: 3.801, h: 7.594, rotate: 180, fill: SAGE });
  txt(s, 'Rebecca', { x: 2.244, y: 2.603, w: 5.542, h: 1.447, rotate: 270, fontSize: 80, bold: true, color: WHITE });
  rect(s, { x: 5.801, y: 4.855, w: 7.532, h: 2.645, fill: PLUM });
  txt(s, 'Photography | Flower | Season | Ingredient | Tips & Trick', { x: 7.468, y: 6.38, w: 5.311, h: 0.452, fontSize: 12, bold: true, italic: true, color: WHITE, lineSpacing: 25, paraSpaceBefore: 12 });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation iriure dolor', { x: 0.928, y: 0.947, w: 2.291, h: 1.784, fontSize: 18, italic: true, color: WHITE, lineSpacing: 24, paraSpaceBefore: 12 });
  note(s, { x: 9.174, y: 3.568, w: 3.103, h: 0.662, labelW: 1.711, label: 'Magna Volutpat', body: 'Mirum est notare quam littera gothica,  nam quam nunc putamus parum claram' });
}

// Slide 2 — Hello & Welcome
function slide2(s) {
  rect(s, { x: 0.388, y: 0.327, w: 12.568, h: 6.863, fill: PLUM });
  circle(s, { x: 4.576, y: 2.009, w: 4.192, h: 4.192, fill: MOSS });
  txt(s, 'Hello &\nWelcome', { x: 0.948, y: 2.573, w: 5.616, h: 1.668, fontSize: 66, bold: true, color: WHITE, lineSpacing: 55 });
  txt(s, 'PLACEHOLDER' +
    'consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et ' +
    'iusto odio. Duis autem vel eum iriure dolor in hendrerit in vulputate velit esse ' +
    'molestie consequat, vel illum dolore eu feugiat nulla', { x: 7.445, y: 4.105, w: 4.671, h: 1.503, fontSize: 10, color: WHITE, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T1, { x: 0.938, y: 4.105, w: 5.26, h: 1.111, fontSize: 18, italic: true, color: WHITE, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 9.323, 0.664, WHITE);
}

// Slide 3 — SHINE
function slide3(s) {
  rect(s, { x: 0, y: 0, w: 9.854, h: 5.488, fill: MIST });
  txt(s, 'SHINE', { x: 2.197, y: 1.616, w: 6.112, h: 2.205, rotate: 270, fontSize: 125, bold: true });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu ' +
    'feugiat nulla', { x: 7.046, y: 2.361, w: 1.992, h: 2.308, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  note(s, { x: 10.608, y: 4.768, w: 1.99, h: 1.223, labelW: 1.711, label: 'Magna Volutpat', body: T8 });
  txt(s, T1, { x: 1.203, y: 1.875, w: 1.509, h: 2.794, fontSize: 14, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 0.53, 0.664);
}

// Slide 4 — quote + gallery
function slide4(s) {
  rect(s, { x: 5.117, y: 0, w: 8.216, h: 6.875, fill: MIST });
  note(s, { x: 8.505, y: 2.76, w: 3.657, h: 0.662, labelW: 1.711, label: 'Magna Volutpat', body: T8 });
  txt(s, T9, { x: 6.039, y: 5.862, w: 2.196, h: 0.662, fontSize: 9, align: 'right', lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Magna Volutpat', { x: 6.529, y: 5.53, w: 1.706, h: 0.327, fontSize: 11, bold: true, align: 'right', lineSpacing: 18, paraSpaceBefore: 12 });
  note(s, { x: 5.673, y: 0.986, w: 1.877, h: 0.902, labelW: 1.711, label: 'Magna Volutpat', body: T10 });
  txt(s, T11, { x: 0.809, y: 1.531, w: 2.578, h: 2.457, fontSize: 20, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  txt(s, 'Kentetsu Takamori', { x: 0.809, y: 3.988, w: 2.338, h: 0.452, fontSize: 12, bold: true, italic: true, lineSpacing: 25, paraSpaceBefore: 12 });
  disc(s, { x: 3.367, y: 4.679, w: 1.75, h: 3.496, rotate: 180, fill: SAGE });
  brand(s, 0.53, 0.664);
}

// Slide 5 — Early
function slide5(s) {
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu ' +
    'feugiat nulla facilisis at.', { x: 10.132, y: 1.872, w: 2.223, h: 2.064, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  note(s, { x: 10.132, y: 5.5, w: 1.711, h: 1.223, labelW: 1.711, label: 'Magna Volutpat', body: T12 });
  rect(s, { x: 0, y: 0, w: 3.308, h: 7.5, fill: PLUM });
  txt(s, 'Early', { x: 1.281, y: 2.162, w: 4.88, h: 2.053, fontSize: 116, bold: true, color: WHITE });
  txt(s, T1, { x: 0.981, y: 0.936, w: 4.644, h: 0.774, fontSize: 14, italic: true, color: WHITE, lineSpacing: 24, paraSpaceBefore: 12 });
  txt(s, 'Mirum est notare quam littera gothica, quam nunc putamus parum claram', { x: 0.855, y: 5.412, w: 1.99, h: 0.942, fontSize: 9, color: WHITE, align: 'right', lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Magna Volutpat', { x: 1.134, y: 5.08, w: 1.711, h: 0.327, fontSize: 11, bold: true, color: WHITE, align: 'right', lineSpacing: 18, paraSpaceBefore: 12 });
  brand(s, 9.323, 0.664);
}

// Slide 6 — GIVE
function slide6(s) {
  rect(s, { x: 0.388, y: 0.327, w: 12.568, h: 6.863, fill: PLUM });
  disc(s, { x: 0.388, y: 1.622, w: 1.75, h: 3.496, fill: SAGE });
  txt(s, 'GIVE', { x: 1.753, y: 2.023, w: 7.807, h: 3.45, rotate: 270, fontSize: 200, bold: true, color: WHITE, align: 'center' });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu ' +
    'PLACEHOLDER' +
    'consequat.', { x: 1.182, y: 4.108, w: 2.618, h: 2.345, fontSize: 10, color: WHITE, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T1, { x: 7.512, y: 3.59, w: 4.719, h: 0.774, fontSize: 14, italic: true, color: WHITE, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 9.323, 0.664, WHITE);
  rect(s, { x: 6.021, y: 4.747, w: 7.312, h: 2.753, fill: PLUM, transparency: 40 });
}

// Slide 7 — Flow
function slide7(s) {
  txt(s, T11, { x: 5.981, y: 4.74, w: 2.919, h: 1.819, fontSize: 16, italic: true, lineSpacing: 25, paraSpaceBefore: 12 });
  txt(s, 'Kentetsu Takamori', { x: 5.981, y: 6.559, w: 2.369, h: 0.403, fontSize: 12, bold: true, lineSpacing: 25, paraSpaceBefore: 12 });
  txt(s, 'PLACEHOLDER' +
    'consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et ' +
    'iusto odio. ', { x: 1.452, y: 3.704, w: 3.01, h: 1.503, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  disc(s, { x: 8.174, y: 2.004, w: 0.99, h: 1.979, rotate: 180, fill: SAGE });
  txt(s, 'Flow', { x: 4.463, y: 0.427, w: 4.888, h: 2.053, fontSize: 116, bold: true });
  disc(s, { x: 0, y: 4.74, w: 1.774, h: 3.543, fill: MIST });
  brand(s, 0.53, 0.664);
}

// Slide 8 — 1987
function slide8(s) {
  rect(s, { x: 6.905, y: 0, w: 6.428, h: 7.5, fill: PLUM, transparency: 10 });
  disc(s, { x: 6.905, y: 1.206, w: 2.547, h: 5.088, fill: MIST });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu ' +
    'feugiat nulla facilisis at vero eros et accumsan et', { x: 9.8, y: 4.831, w: 2.474, h: 2.027, fontSize: 10, color: WHITE, lineSpacing: 20, paraSpaceBefore: 12 });
  note(s, { x: 10.281, y: 2.801, w: 1.869, h: 0.942, color: WHITE, labelW: 1.708, label: 'Magna Volutpat', body: 'Nam liber tempor cum soluta nobis eleifend option congue' });
  note(s, { x: 7.71, y: 1.179, w: 1.877, h: 0.942, color: WHITE, labelW: 1.711, label: 'Magna Volutpat', body: T10 });
  txt(s, '1987', { x: 0.94, y: 3.551, w: 8.066, h: 4.14, fontSize: 240, bold: true, color: WHITE, align: 'center' });
  txt(s, T1, { x: 1.112, y: 2.776, w: 4.719, h: 0.774, fontSize: 14, italic: true, color: WHITE, lineSpacing: 24, paraSpaceBefore: 12 });
  note(s, { x: 10.281, y: 1.206, w: 1.998, h: 0.902, color: WHITE, labelW: 1.711, label: 'Magna Volutpat', body: T12 });
  brand(s, 0.53, 0.664, WHITE);
}

// Slide 9 — Orchid
function slide9(s) {
  rect(s, { x: 0.453, y: 3.193, w: 8.775, h: 2.807, fill: MIST });
  txt(s, 'Orchid', { x: 0.843, y: 2.745, w: 4.244, h: 1.447, fontSize: 80, bold: true });
  txt(s, T9, { x: 6.587, y: 4.047, w: 2.22, h: 0.662, fontSize: 9, align: 'right', lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Magna Volutpat', { x: 6.413, y: 3.715, w: 2.394, h: 0.353, fontSize: 11, bold: true, align: 'right', lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, T1, { x: 1.16, y: 4.915, w: 4.719, h: 0.736, fontSize: 14, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  txt(s, T13, { x: 6.963, y: 1.121, w: 5.341, h: 0.942, fontSize: 18, italic: true, lineSpacing: 30, paraSpaceBefore: 12 });
  txt(s, 'Maya Angelou', { x: 6.97, y: 2.021, w: 2.451, h: 0.408, fontSize: 14, bold: true, italic: true, lineSpacing: 25, paraSpaceBefore: 12 });
  brand(s, 0.53, 0.664);
}

// Slide 10 — NEVER WILT / TILL END
function slide10(s) {
  disc(s, { x: 2.115, y: 4.798, w: 1.283, h: 2.563, rotate: 90, fill: MAUVE });
  txt(s, 'TILL END', { x: 4.66, y: 4.916, w: 8.107, h: 2.036, fontSize: 115, bold: true, color: WHITE, align: 'center', wrap: false });
  txt(s, 'NEVER WILT', { x: 0.253, y: 0.354, w: 11.568, h: 2.036, fontSize: 115, bold: true, color: WHITE, align: 'center' });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation duis autem', { x: 2.513, y: 6.178, w: 5.037, h: 0.774, fontSize: 16, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 0.53, 0.664);
}

// Slide 11 — two info cards
function slide11(s) {
  rect(s, { x: 6.958, y: 4.615, w: 3.382, h: 1.549, fill: PLUM, transparency: 30 });
  txt(s, 'Magna Volutpat', { x: 7.842, y: 4.908, w: 1.528, h: 0.323, fontSize: 10, bold: true, color: WHITE, lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, '15/3', { x: 7.359, y: 4.908, w: 0.598, h: 0.323, fontSize: 10, bold: true, color: WHITE, lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, T4, { x: 7.359, y: 5.186, w: 2.8, h: 0.662, fontSize: 9, color: WHITE, lineSpacing: 20, paraSpaceBefore: 12 });
  rect(s, { x: 0.763, y: 4.615, w: 4.283, h: 1.549, fill: MIST });
  txt(s, 'Magna Volutpat', { x: 1.875, y: 4.878, w: 1.528, h: 0.353, fontSize: 10, bold: true, lineSpacing: 18, paraSpaceBefore: 12 });
  note(s, { x: 1.393, y: 4.878, w: 3.286, h: 0.662, labelW: 0.598, labelSize: 10, gap: 0.355, label: '15/3', body: T14 });
  txt(s, T2, { x: 2.553, y: 6.501, w: 8.228, h: 0.438, fontSize: 14, italic: true, align: 'center', lineSpacing: 24, paraSpaceBefore: 12 });
  txt(s, T15, { x: 0.763, y: 2.542, w: 5.115, h: 1.503, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T2, { x: 0.763, y: 1.499, w: 5.115, h: 0.774, fontSize: 18, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  disc(s, { x: 11.44, y: 3.57, w: 1.299, h: 2.594, fill: MIST });
  brand(s, 0.53, 0.664);
}

// Slide 12 — Bucket
function slide12(s) {
  rect(s, { x: -0.069, y: 0, w: 7.967, h: 7.5, fill: MIST });
  disc(s, { x: 6.542, y: 3.118, w: 1.414, h: 2.824, rotate: 180, fill: MOSS });
  rect(s, { x: 1.978, y: 3.714, w: 4.225, h: 1.632, fill: PLUM });
  note(s, { x: 2.419, y: 3.841, w: 3.512, h: 0.942, color: WHITE, labelW: 1.528, labelSize: 10, gap: 0.357, label: 'Magna Volutpat', body: 'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie, vel illum dolore' });
  txt(s, 'Bucket', { x: 3.233, y: 5.706, w: 4.136, h: 1.313, fontSize: 72, bold: true });
  txt(s, T2, { x: 0.763, y: 2.096, w: 6.077, h: 0.774, fontSize: 18, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  txt(s, 'Magna Volutpat', { x: 8.762, y: 0.588, w: 1.528, h: 0.353, fontSize: 10, bold: true, lineSpacing: 18, paraSpaceBefore: 12 });
  note(s, { x: 8.279, y: 0.588, w: 3.286, h: 0.662, labelW: 0.598, labelSize: 10, gap: 0.355, label: '15/3', body: T14 });
  brand(s, 0.53, 0.664);
}

// Slide 13 — TOUCHES
function slide13(s) {
  disc(s, { x: 7.028, y: 2.75, w: 2.627, h: 5.247, fill: MIST });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu ' +
    'feugiat nulla facilisis at vero eros et accumsan et iusto odio. Duis autem vel eum ' +
    'iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel', { x: 1.2, y: 4.087, w: 3.776, h: 2.027, fontSize: 10, align: 'right', lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T16, { x: 1.763, y: 1.51, w: 3.214, h: 1.784, fontSize: 14, italic: true, align: 'right', lineSpacing: 24, paraSpaceBefore: 12 });
  txt(s, 'TOUCHES', { x: 2.284, y: 2.441, w: 8.685, h: 1.952, rotate: 270, fontSize: 110, bold: true });
  note(s, { x: 7.603, y: 4.087, w: 3.005, h: 0.902, labelW: 2.507, label: 'Magna Volutpat', body: T3 });
  brand(s, 9.323, 0.664);
}

// Slide 14 — split panel
function slide14(s) {
  rect(s, { x: 0, y: 0, w: 5.836, h: 7.5, fill: PLUM });
  note(s, { x: 0.895, y: 5.587, w: 3.8, h: 0.662, color: WHITE, labelW: 1.706, label: 'Magna Volutpat', body: T3 });
  note(s, { x: 7.938, y: 5.56, w: 2.842, h: 0.942, labelW: 1.77, label: 'Magna Volutpat', body: 'Duis autem vel eum iriure dolor in hendrerit in vulputate velit esse molestie, vel ' +
    'illum dolore nulla' });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu ' +
    'feugiat nulla facilisis at vero eros et accumsan et iusto', { x: 7.226, y: 2.719, w: 5.445, h: 0.942, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T16, { x: 7.226, y: 1.303, w: 5.321, h: 1.111, fontSize: 14, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 0.53, 0.664, WHITE);
}

// Slide 15 — SOFT
function slide15(s) {
  disc(s, { x: 11.964, y: 4.589, w: 1.37, h: 2.737, rotate: 90, fill: MIST });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation iriure dolor in hendrerit ' +
    'nam odio moliestie mirum lorem ipsum', { x: 3.491, y: 3.75, w: 8.263, h: 0.774, fontSize: 18, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  txt(s, 'SOFT', { x: -2.914, y: 2.157, w: 7.763, h: 3.45, rotate: 270, fontSize: 199, bold: true, color: WHITE });
  note(s, { x: 8.154, y: 5.272, w: 3.899, h: 0.942, labelW: 1.706, label: 'Magna Volutpat', body: T17 });
  brand(s, 9.323, 0.664, WHITE);
  rect(s, { x: 3.491, y: 0, w: 9.842, h: 2.648, fill: PLUM, transparency: 30 });
}

// Slide 16 — Fragrant
function slide16(s) {
  rect(s, { x: 8.683, y: 3.317, w: 3, h: 4.183, fill: PLUM });
  disc(s, { x: -0.275, y: 0.856, w: 2.714, h: 5.422, rotate: 180, fill: MIST });
  txt(s, 'Fragrant', { x: 0.379, y: 1.988, w: 6.588, h: 1.717, fontSize: 96, bold: true });
  note(s, { x: 9.233, y: 4.935, w: 1.706, h: 1.503, color: WHITE, labelW: 1.706, label: 'Magna Volutpat', body: 'Duis autem vel eum iriure dolor in hendrerit in velit esse molestie consequat, vel ' +
    'illum dolore eros et' });
  txt(s, T18, { x: 5.78, y: 0.852, w: 6.563, h: 0.942, fontSize: 11, lineSpacing: 20, paraSpaceBefore: 12 });
  brand(s, 0.53, 0.664);
}

// Slide 17 — Growth
function slide17(s) {
  rect(s, { x: 0, y: 4.925, w: 5.955, h: 2.575, fill: MIST });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu ' +
    'feugiat nulla facilisis at vero eros et', { x: 1.166, y: 5.601, w: 3.923, h: 1.223, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T2, { x: 1.166, y: 1.191, w: 3.22, h: 1.111, fontSize: 14, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  disc(s, { x: 5.955, y: 1.836, w: 1.547, h: 3.09, fill: MOSS });
  txt(s, 'Growth', { x: 2.439, y: 2.94, w: 5.696, h: 1.717, fontSize: 96, bold: true, color: WHITE });
  brand(s, 9.323, 0.664);
}

// Slide 18 — FINE
function slide18(s) {
  disc(s, { x: 0, y: 2.604, w: 2.578, h: 5.151, fill: MIST });
  rect(s, { x: 7.533, y: 0, w: 5.801, h: 7.5, fill: PLUM });
  txt(s, T18, { x: 1.304, y: 4.084, w: 3.554, h: 1.466, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'FINE', { x: 3.924, y: 2.772, w: 5.788, h: 2.625, rotate: 270, fontSize: 150, bold: true, color: WHITE });
  txt(s, T1, { x: 1.289, y: 1.422, w: 2.309, h: 1.784, fontSize: 14, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 0.53, 0.664);
}

// Slide 19 — quote grid
function slide19(s) {
  rect(s, { x: -0, y: 3.739, w: 8.221, h: 3.761, fill: MIST });
  rect(s, { x: 3.594, y: 0, w: 5.826, h: 3.739, fill: PLUM });
  txt(s, T13, { x: 4.08, y: 1.947, w: 4.848, h: 0.897, fontSize: 18, italic: true, color: WHITE, lineSpacing: 30, paraSpaceBefore: 12 });
  txt(s, 'Maya Angelou', { x: 4.087, y: 2.847, w: 2.451, h: 0.408, fontSize: 14, bold: true, italic: true, color: WHITE, lineSpacing: 25, paraSpaceBefore: 12 });
  txt(s, 'Duis autem vel eum iriure dolor in hendrerit in velit esse molestie consequat, vel ' +
    'illum dolore', { x: 9.906, y: 0.991, w: 2.564, h: 0.942, fontSize: 9, align: 'right', lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Magna Volutpat', { x: 10.764, y: 0.659, w: 1.706, h: 0.327, fontSize: 11, bold: true, align: 'right', lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, T15, { x: 0.839, y: 5.429, w: 6.153, h: 1.223, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Hendrerit in velit esse molestie vel illum dolore eros', { x: 10.141, y: 2.459, w: 2.329, h: 0.662, fontSize: 9, align: 'right', lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Magna Volutpat', { x: 10.764, y: 2.127, w: 1.706, h: 0.327, fontSize: 11, bold: true, align: 'right', lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, T2, { x: 0.839, y: 4.291, w: 6.153, h: 0.774, fontSize: 18, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 0.53, 0.664);
  rect(s, { x: 8.221, y: 3.739, w: 5.113, h: 2.913, fill: PLUM, transparency: 30 });
}

// Slide 20 — National Flower
function slide20(s) {
  rect(s, { x: 0, y: -0.062, w: 12.54, h: 5.562, fill: MIST });
  disc(s, { x: 0.982, y: 4.524, w: 1.969, h: 3.933, rotate: 90, fill: MOSS });
  note(s, { x: 3.855, y: 5.694, w: 1.852, h: 0.942, labelW: 1.706, gap: 0.302, label: 'Magna Volutpat', body: T4 });
  note(s, { x: 6.185, y: 5.694, w: 1.852, h: 0.942, labelW: 1.706, gap: 0.302, label: 'Magna Volutpat', body: T19 });
  note(s, { x: 8.436, y: 5.694, w: 1.852, h: 0.942, labelW: 1.706, gap: 0.302, label: 'Magna Volutpat', body: 'PLACEHOLDER' });
  note(s, { x: 10.688, y: 5.694, w: 1.852, h: 0.902, labelW: 1.706, gap: 0.302, label: 'Magna Volutpat', body: 'Velit esse molestie consequat, vel illum dolore eu feugiat' });
  txt(s, 'National Flower', { x: 0.604, y: 0.771, w: 7.432, h: 0.825, fontSize: 60, bold: true, lineSpacing: 50 });
  txt(s, T2, { x: 0.851, y: 2.979, w: 1.674, h: 2.121, fontSize: 14, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 9.323, 0.664);
}

// Slide 21 — Sparkling Blossom
function slide21(s) {
  rect(s, { x: 0, y: 3, w: 9.203, h: 3.275, fill: MIST });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation duis autem vel eum iriure ' +
    'dolor in nam vulputate velit esse molestie consequat, vel illum dolore eu feugiat ' +
    'nulla facilisis', { x: 0.798, y: 5.253, w: 7.942, h: 0.662, fontSize: 11, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation iriure', { x: 10.01, y: 3, w: 2.195, h: 1.447, fontSize: 14, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  txt(s, 'Sparkling\nBlossom', { x: 7.387, y: 0.867, w: 4.439, h: 1.616, fontSize: 48, bold: true, lineSpacing: 54 });
  note(s, { x: 0.798, y: 1.638, w: 2.526, h: 0.662, labelW: 1.706, label: 'Magna Volutpat', body: T4 });
  brand(s, 9.562, 6.53);
}

// Slide 22 — timeline 1982 / 1991 / 2003
function slide22(s) {
  txt(s, T5, { x: 3.221, y: 1.469, w: 7.382, h: 0.662, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'PLACEHOLDER' +
    'consequat, vel illum dolore eu feugiat nulla facilisis at vero eros et accumsan et ' +
    'iusto odio', { x: 3.231, y: 2.73, w: 7.37, h: 0.662, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T5, { x: 3.231, y: 4.086, w: 7.37, h: 0.662, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Subtitle', { x: 3.221, y: 1.152, w: 1.706, h: 0.353, fontSize: 14, bold: true, lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, 'Subtitle', { x: 3.221, y: 2.413, w: 1.706, h: 0.353, fontSize: 14, bold: true, lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, 'Subtitle', { x: 3.231, y: 3.769, w: 1.706, h: 0.353, fontSize: 14, bold: true, lineSpacing: 18, paraSpaceBefore: 12 });
  rect(s, { x: 1.914, y: 5.266, w: 5.014, h: 1.632, fill: MIST });
  rect(s, { x: 6.929, y: 5.266, w: 5.286, h: 1.632, fill: PLUM });
  txt(s, 'Magna Volutpat', { x: 7.96, y: 5.349, w: 1.528, h: 0.323, fontSize: 10, bold: true, color: WHITE, lineSpacing: 18, paraSpaceBefore: 12 });
  note(s, { x: 7.477, y: 5.349, w: 4.367, h: 0.942, color: WHITE, labelW: 0.598, labelSize: 10, gap: 0.4, label: '15/3', body: 'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie, vel illum dolore nulla' });
  txt(s, 'Magna Volutpat', { x: 2.741, y: 5.385, w: 1.528, h: 0.353, fontSize: 10, bold: true, lineSpacing: 18, paraSpaceBefore: 12 });
  note(s, { x: 2.258, y: 5.385, w: 4.241, h: 0.942, labelW: 0.598, labelSize: 10, gap: 0.401, label: '15/3', body: T17 });
  txt(s, '1982', { x: 1.797, y: 0.962, w: 1.126, h: 0.572, fontSize: 28, bold: true, paraSpaceBefore: 12 });
  txt(s, '2003', { x: 1.797, y: 3.566, w: 1.203, h: 0.572, fontSize: 28, bold: true, paraSpaceBefore: 12 });
  txt(s, '1991', { x: 1.797, y: 2.24, w: 1.126, h: 0.572, fontSize: 28, bold: true, paraSpaceBefore: 12 });
  brand(s, 9.323, 0.664);
}

// Slide 23 — Break slide
function slide23(s) {
  disc(s, { x: 11.461, y: -1.313, w: 2.44, h: 4.874, rotate: 180, fill: MOSS });
  disc(s, { x: 3.7, y: 2.659, w: 3.16, h: 6.312, fill: MIST });
  txt(s, 'Break\t\t\t\t\t\tslide', { x: -0.475, y: 2.27, w: 14.266, h: 1.952, fontSize: 110, bold: true, align: 'center' });
  txt(s, T5, { x: 6.236, y: 3.872, w: 2.264, h: 2.064, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T6, { x: 6.236, y: 1.546, w: 2.421, h: 1.447, fontSize: 18, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 0.53, 0.664);
}

// Slide 24 — BREAK slide
function slide24(s) {
  rect(s, { x: 0.966, y: -0.062, w: 10.239, h: 6.175, fill: MIST });
  disc(s, { x: 11.206, y: 4.025, w: 1.84, h: 3.677, fill: MOSS });
  txt(s, 'BREAK', { x: 2.517, y: 5.441, w: 8.835, h: 0.942, fontSize: 166, bold: true, lineSpacing: 60 });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu ' +
    'feugiat nulla facilisis', { x: 2.316, y: 6.367, w: 7.54, h: 0.662, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'slide', { x: 9.796, y: 6.236, w: 2.072, h: 0.841, fontSize: 44, bold: true, align: 'right' });
  brand(s, 0.53, 0.664);
}

// Slide 25 — Breakslide
function slide25(s) {
  rect(s, { x: 0.659, y: 4.504, w: 12.029, h: 2.408, fill: PLUM, transparency: 25 });
  note(s, { x: 6.931, y: 5.121, w: 4.549, h: 0.662, color: WHITE, labelW: 1.706, label: 'Magna Volutpat', body: 'Nam liber tempor cum soluta nobis eleifend option congue nihil. Mirum est notare ' +
    'quam littera gothica, quam nunc putamus' });
  txt(s, 'Breakslide', { x: 2.468, y: 3.106, w: 6.3, h: 1.313, rotate: 270, fontSize: 72, bold: true, color: WHITE, align: 'center' });
  txt(s, T6, { x: 2.014, y: 3.742, w: 2.421, h: 1.447, fontSize: 18, italic: true, color: WHITE, lineSpacing: 24, paraSpaceBefore: 12 });
  disc(s, { x: 6.292, y: -0.612, w: 1.492, h: 2.981, fill: MIST });
  brand(s, 9.323, 0.664, WHITE);
}

// Slide 26 — Android mock-up
function slide26(s) {
  phoneMock(s, { x: 8.292, y: 1.778, w: 2.083, h: 4.222, radius: 0.1,
    bezel: 0.097, screen: 'A6ECE4' });
  disc(s, { x: 2.814, y: -0.904, w: 1.054, h: 2.106, rotate: 180, fill: PLUM });
  disc(s, { x: -0.288, y: -1.607, w: 3.16, h: 6.312, fill: MIST });
  txt(s, T20, { x: 1.167, y: 5.144, w: 6.08, h: 0.905, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T2, { x: 1.167, y: 3.594, w: 4.391, h: 1.111, fontSize: 18, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  txt(s, T19, { x: 5.011, y: 2.055, w: 2.885, h: 0.662, fontSize: 9, align: 'right', lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Magna Volutpat', { x: 6.19, y: 1.753, w: 1.706, h: 0.327, fontSize: 11, bold: true, align: 'right', lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, 'Android', { x: 8.984, y: 1.923, w: 4.995, h: 1.447, rotate: 270, fontSize: 80, bold: true });
  brand(s, 0.53, 0.664);
}

// Slide 27 — iphone mock-up
function slide27(s) {
  disc(s, { x: 2.341, y: 3.823, w: 1.84, h: 3.677, rotate: 180, fill: MOSS });
  rect(s, { x: 4.301, y: -0.062, w: 9.032, h: 6.035, fill: MIST });
  phoneMock(s, { x: 3.486, y: 1.708, w: 2.222, h: 4.764, radius: 0.12,
    bezel: 0.07, notch: true, screen: '7CE9E0' });
  txt(s, T20, { x: 6.716, y: 3.696, w: 3.336, h: 1.503, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, T3, { x: 0.904, y: 2.65, w: 1.82, h: 1.503, fontSize: 9, align: 'right', lineSpacing: 20, paraSpaceBefore: 12 });
  txt(s, 'Magna Volutpat', { x: 1.018, y: 2.318, w: 1.706, h: 0.327, fontSize: 11, bold: true, align: 'right', lineSpacing: 18, paraSpaceBefore: 12 });
  txt(s, 'iphone', { x: 9.953, y: 2.569, w: 4.434, h: 1.447, rotate: 270, fontSize: 80, bold: true });
  txt(s, T6, { x: 6.793, y: 1.013, w: 4.271, h: 0.774, fontSize: 18, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 0.53, 0.664);
}

// Slide 28 — Desktop mock-up
function slide28(s) {
  disc(s, { x: 3.533, y: -1.208, w: 2.94, h: 5.874, rotate: 180, fill: MIST });
  monitorMock(s, { x: 5.042, y: 2.361, w: 6.472, h: 3.903, screen: 'A6ECE4' });
  txt(s, 'Desktop', { x: 1.015, y: 0.668, w: 5.446, h: 1.447, fontSize: 80, bold: true });
  txt(s, 'Ut wisi enim ad minim veniam, quis nostrud exerci tation Duis autem vel eum iriure ' +
    'dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu ' +
    'feugiat nulla facilisis at vero eros et accumsan', { x: 1.546, y: 3.682, w: 2.218, h: 2.345, fontSize: 10, lineSpacing: 20, paraSpaceBefore: 12 });
  disc(s, { x: 11.522, y: 4.437, w: 1.054, h: 2.106, fill: PLUM });
  txt(s, T7, { x: 6.929, y: 1.005, w: 5.321, h: 0.736, fontSize: 14, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  brand(s, 0.53, 6.589);
}

// Slide 29 — WORLD map
function slide29(s) {
  disc(s, { x: 4.003, y: 5.02, w: 1.656, h: 3.308, rotate: 270, fill: MIST });
  disc(s, { x: -0.003, y: 2.941, w: 2.94, h: 5.874, fill: MOSS });
  worldMap(s, { x: 5.592, y: 0.884, w: 6.842, h: 3.86 });
  note(s, { x: 1.9, y: 3.086, w: 2.045, h: 1.223, labelW: 1.706, label: 'Magna Volutpat', body: T3 });
  note(s, { x: 1.9, y: 5.351, w: 2.134, h: 1.223, labelW: 1.77, label: 'Magna Volutpat', body: T21 });
  note(s, { x: 4.779, y: 5.364, w: 2.045, h: 1.223, labelW: 1.706, label: 'Magna Volutpat', body: T3 });
  note(s, { x: 7.656, y: 5.401, w: 2.134, h: 1.223, labelW: 1.77, label: 'Magna Volutpat', body: T21 });
  txt(s, 'WORLD', { x: 0.764, y: 1.279, w: 4.081, h: 1.111, fontSize: 60, bold: true });
  note(s, { x: 10.434, y: 5.362, w: 2.045, h: 1.223, labelW: 1.706, label: 'Magna Volutpat', body: T3 });
  brand(s, 0.53, 0.664);
}

// Slide 30 — Table
function slide30(s) {
  disc(s, { x: 11.045, y: 3.334, w: 2.288, h: 4.571, rotate: 180, fill: MIST });
  dataTable(s, { x: 0.93, y: 1.774, w: 9.207, h: 4.809 }, [
    ['Name', 'Claritas', 'Laoreet', 'Vero', 'Eros'],
    ['Kirito', 'Glass', 'Polycarbonat', 'Plastic', 'Polycarbonat'],
    ['Asuna', '77%', '-', '51%', '38%'],
    ['Hikigaya', 'High', '-', 'Low', 'Low'],
    ['Megumin', 'yes', 'yes', '-', '-'],
    ['Kazuma', 'Snapdragon', 'Oxygen', '-', 'Snapdragon'],
    ['Lalatina', '+21', '+54', '+86', '+13'],
    ['Aqua', 'Red Velvet', 'Blue Ocean', 'Tosca Green', 'Light Brown'],
  ]);
  txt(s, 'Table', { x: 10.411, y: 2.827, w: 3.75, h: 1.447, rotate: 270, fontSize: 80, bold: true });
  txt(s, T7, { x: 0.93, y: 1.142, w: 9.32, h: 0.438, fontSize: 14, italic: true, lineSpacing: 24, paraSpaceBefore: 12 });
  disc(s, { x: 12.279, y: -0.43, w: 1.054, h: 2.106, rotate: 180, fill: PLUM });
  brand(s, 9.323, 0.664);
}

// Slide 31 — sky quote
function slide31(s) {
  rect(s, { x: 3.024, y: 3.887, w: 10.309, h: 3.613, fill: MIST });
  disc(s, { x: -0.448, y: -0.562, w: 3.472, h: 6.936, rotate: 180, fill: PLUM });
  txt(s, 'The blue of the sky is one of the most special colors in the world, because the ' +
    'color is deep but see-through both at the same time.', { x: 6.667, y: 4.708, w: 5.536, h: 1.363, fontSize: 16, italic: true, lineSpacing: 30, paraSpaceBefore: 12 });
  txt(s, 'Cynthia Kadohata', { x: 6.667, y: 6.113, w: 2.451, h: 0.408, fontSize: 14, bold: true, italic: true, lineSpacing: 25, paraSpaceBefore: 12 });
  rect(s, { x: 3.767, y: 2.255, w: 5.841, h: 1.632, fill: MOSS, transparency: 20 });
  txt(s, 'Magna Volutpat', { x: 4.798, y: 2.526, w: 1.528, h: 0.323, fontSize: 10, bold: true, color: WHITE, lineSpacing: 18, paraSpaceBefore: 12 });
  note(s, { x: 4.316, y: 2.526, w: 4.858, h: 0.662, color: WHITE, labelW: 0.598, labelSize: 10, gap: 0.401, label: '15/3', body: 'Ut wisi enim ad minim veniam, quis nostrud exerci tation  vel eum iriure dolor in ' +
    'hendrerit in vulputate velit esse molestie, vel illum dolore nulla' });
  brand(s, 9.323, 0.664);
}

// Slide 32 — Get it Touch
function slide32(s) {
  rect(s, { x: 5.092, y: 1.014, w: 8.242, h: 6.486, fill: MIST });
  disc(s, { x: 1.329, y: 2.327, w: 3.461, h: 6.914, rotate: 270, fill: MOSS });
  txt(s, 'Address', { x: 7.973, y: 1.617, w: 1.545, h: 0.522, fontSize: 14, bold: true, lineSpacing: 30 });
  txt(s, '21 Jump Street, 2345, Jakarta, Indonesia', { x: 7.973, y: 2.051, w: 3.603, h: 0.353, fontSize: 10, lineSpacing: 18 });
  txt(s, '(Office Hours) Monday – Friday, 08.00 – 16.00', { x: 7.973, y: 2.34, w: 3.603, h: 0.353, fontSize: 10, lineSpacing: 18 });
  txt(s, 'Get It Touch', { x: 7.973, y: 2.934, w: 2.011, h: 0.522, fontSize: 14, bold: true, lineSpacing: 30 });
  txt(s, 'scaramouse@mail.com', { x: 7.973, y: 3.369, w: 3.482, h: 0.353, fontSize: 10, lineSpacing: 18 });
  txt(s, '(+62) 856-0987-7832', { x: 7.973, y: 3.7, w: 3.482, h: 0.353, fontSize: 10, lineSpacing: 18 });
  txt(s, '(0645) 7987-9293', { x: 7.973, y: 4.016, w: 3.482, h: 0.353, fontSize: 10, lineSpacing: 18 });
  txt(s, 'scaramouse', { x: 8.51, y: 4.663, w: 1.76, h: 0.353, fontSize: 10, lineSpacing: 18 });
  txt(s, 'Scaramouse_id', { x: 8.51, y: 5.182, w: 1.76, h: 0.353, fontSize: 10, lineSpacing: 18 });
  txt(s, 'scaramouse-_id', { x: 8.491, y: 5.744, w: 1.76, h: 0.353, fontSize: 10, lineSpacing: 18 });
  txt(s, 'www.scaramouse.com', { x: 8.491, y: 6.263, w: 1.937, h: 0.353, fontSize: 10, lineSpacing: 18 });
  txt(s, 'Get it Touch', { x: 1.532, y: 3.58, w: 6.496, h: 1.363, rotate: 270, fontSize: 66, bold: true, align: 'center', lineSpacing: 90 });
  iconDot(s, 8.074, 4.67, 'f');   // facebook
  iconDot(s, 8.074, 5.194, 'o');  // instagram
  iconDot(s, 8.096, 5.751, 'p');  // pinterest
  iconDot(s, 8.111, 6.275, 'w');  // website
  brand(s, 0.53, 0.664);
}

// Slide 33 — Thank You
function slide33(s) {
  rect(s, { x: 0.659, y: 0.612, w: 12.029, h: 6.3, fill: PLUM, transparency: 30 });
  txt(s, 'Thank You', { x: 1.251, y: 4.29, w: 5.515, h: 1.212, fontSize: 66, bold: true, color: WHITE, align: 'center' });
  txt(s, T7, { x: 1.251, y: 5.502, w: 5.321, h: 0.736, fontSize: 14, italic: true, color: WHITE, lineSpacing: 24, paraSpaceBefore: 12 });
  disc(s, { x: 10.346, y: 3.02, w: 2.989, h: 5.971, rotate: 270, fill: MIST });
}


// ---------------------------------------------------------------- table

/**
 * Slide 30's data table: dark header row, transparent body rows separated by
 * thin white rules.
 */
function dataTable(s, o, rows) {
  const head = { fill: { color: PLUM }, color: WHITE, fontSize: 12,
    border: [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }] };
  const body = { color: '000000', fontSize: 10,
    border: [{ type: 'none' }, { type: 'none' },
             { type: 'solid', color: WHITE, pt: 1 }, { type: 'none' }] };
  s.addTable(
    rows.map((row, r) => row.map((cell) => ({ text: cell, options: r === 0 ? head : body }))),
    { x: o.x, y: o.y, w: o.w, colW: Array(rows[0].length).fill(o.w / rows[0].length),
      rowH: [0.588].concat(Array(rows.length - 1).fill(0.603)),
      align: 'center', valign: 'middle', fontFace: FONT, margin: 0 }
  );
}

// ------------------------------------------------------------------ deck

const SLIDES = [
  slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9,
  slide10, slide11, slide12, slide13, slide14, slide15, slide16, slide17,
  slide18, slide19, slide20, slide21, slide22, slide23, slide24, slide25,
  slide26, slide27, slide28, slide29, slide30, slide31, slide32, slide33,
];

// Slide 10 is the only one that sits on the light background.
const BACKGROUNDS = { 10: MIST };

function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.3333333, height: 7.5 }); // 12192000 x 6858000 EMU
  pptx.layout = 'WIDE';
  pptx.title = 'Rebecca Florist — Presentation Template';

  SLIDES.forEach((draw, i) => {
    const s = pptx.addSlide();
    s.background = { color: BACKGROUNDS[i + 1] || MAUVE };
    draw(s);
  });

  return pptx.writeFile({
    fileName: path.join(__dirname, '18056061-cca1-443d-91c2-4c391405ec9e_grok_final.pptx'),
  });
}

build().then((f) => console.log('wrote', f));
