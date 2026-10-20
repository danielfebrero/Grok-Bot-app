/**
 * "Celebrity" deck - rebuilt with pptxgenjs.
 *
 * 32 slides, 13.333 x 7.5 in (16:9).  Typography is Oswald for headings /
 * labels and Montserrat for body copy, on a black-and-white palette.
 * Raster photos and icon artwork from the original file are re-drawn here as
 * native shapes / labelled placeholder rectangles.
 */

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const INK = '0D0D0D'; // near-black used for headings, bars, cards
const MID = '404040'; // secondary dark grey (alternating cards)
const BODY = '595959'; // body copy grey
const RULE = 'D9D9D9'; // light grey rules / map silhouettes
const WHITE = 'FFFFFF';
const PHOTO = 'E4E4E4'; // fill for photo placeholders

const HEAD = 'Oswald';
const TEXT = 'Montserrat';

/* -------------------------------------------------------------- text styles */

// Small caps-ish section label, e.g. "Description Here"
const LABEL = { fontFace: HEAD, fontSize: 20, charSpacing: 0.5, color: INK, valign: 'top', margin: [14.4, 14.4, 7.2, 7.2] };
// Big slide heading, e.g. "Subtitle Section Here"
const TITLE = { fontFace: HEAD, fontSize: 38, color: INK, valign: 'bottom' };
// Justified paragraph copy with double leading
const PARA = { fontFace: TEXT, fontSize: 11, color: BODY, align: 'justify', lineSpacingMultiple: 2, valign: 'top' };
// White knock-out pill button label
const PILL = { fontFace: HEAD, fontSize: 14, charSpacing: 0.5, color: WHITE, align: 'center', valign: 'middle', margin: [14.4, 14.4, 7.2, 7.2] };

const LOREM = {
  short: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt.',
  labore: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore.',
  dolore: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore.',
  aliqua: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.',
  commodo: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo.',
  consequat: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.',
  nostrud: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis nostrud.',
  colore: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, se eiusmod tempor incididunt ut labore colore.',
  tempor: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, se eiusmod tempor.',
  sedDo: 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor.',
  incididunt: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt labore et dolore magna aliqua.',
  incididuntShort: 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt labore.',
};

/* ------------------------------------------------------------------ helpers */

const at = (x, y, w, h) => ({ x, y, w, h });

/** Section label ("Description Here" style). */
function label(slide, text, x, y, opts = {}) {
  slide.addText(text, Object.assign({}, LABEL, at(x, y, 2.471, 0.539), opts));
}

/** Body paragraph. */
function para(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, Object.assign({}, PARA, at(x, y, w, h), opts));
}

/** Big slide heading; `text` may be an array for a multi-line stack. */
function title(slide, text, x, y, w, h, opts = {}) {
  const lines = [].concat(text);
  const body = lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } }));
  const lead = lines.length > 1 ? { lineSpacingMultiple: 1.2 } : {};
  slide.addText(body, Object.assign({}, TITLE, lead, at(x, y, w, h), opts));
}

/** Solid black pill button with centred Oswald caption. */
function pill(slide, text, x, y, w = 1.871, h = 0.513) {
  slide.addShape('rect', Object.assign(at(x, y, w, h), { fill: { color: INK } }));
  slide.addText(text, Object.assign({}, PILL, at(x, y, w, h)));
}

/** Black disc with a two-line white caption - the "Personal Identity" badge. */
function badge(slide, x, y, d = 1.997, lines = ['Personal', 'Identity']) {
  slide.addShape('ellipse', Object.assign(at(x, y, d, d), { fill: { color: INK }, line: { color: WHITE, width: 12 } }));
  slide.addText(lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } })), Object.assign({}, PILL, {
    fontSize: 20, lineSpacingMultiple: 1.2, valign: 'middle',
  }, at(x + d * 0.081, y + d * 0.25, d * 0.838, d * 0.488)));
}

/** Stand-in for a photograph from the source deck. */
function photo(slide, x, y, w, h, caption = '[image]') {
  slide.addShape('rect', Object.assign(at(x, y, w, h), { fill: { color: PHOTO } }));
  if (caption) {
    slide.addText(caption, Object.assign(at(x, y, w, h), {
      fontFace: TEXT, fontSize: 12, color: '9A9A9A', align: 'center', valign: 'middle',
    }));
  }
}



/** Horizontal progress bar: grey track plus a black fill of `pct` width. */
function progress(slide, caption, pct, x, y, w) {
  slide.addText(caption, Object.assign({}, at(x - 0.119, y - 0.661, 1.206, 0.456), {
    fontFace: HEAD, fontSize: 16, charSpacing: 0.5, color: INK, valign: 'top', lineSpacingMultiple: 1.5,
  }));
  slide.addText(pct + '%', Object.assign({}, at(x + w - 0.72, y - 0.672, 0.72, 0.456), {
    fontFace: HEAD, fontSize: 16, charSpacing: 0.5, color: INK, align: 'right', valign: 'top', lineSpacingMultiple: 1.5,
  }));
  slide.addShape('line', Object.assign(at(x, y, w, 0), { line: { color: RULE, width: 12 } }));
  slide.addShape('line', Object.assign(at(x, y, w * pct / 100, 0), { line: { color: INK, width: 12 } }));
}

/* -------------------------------------------------------------------- icons */

// Each pictogram is a list of [shape, x, y, w, h, role] primitives placed in a
// unit square, so it can be stamped at any size (stands in for the deck's icon
// artwork).  'bg' parts are painted in the backdrop colour to carve holes.
const GLYPHS = {
  influencer: [ // head and shoulders inside two broadcast arcs
    ['ellipse', 0.04, 0.04, 0.92, 0.92, 'fg'],
    ['ellipse', 0.12, 0.12, 0.76, 0.76, 'bg'],
    ['ellipse', 0.19, 0.19, 0.62, 0.62, 'fg'],
    ['ellipse', 0.27, 0.27, 0.46, 0.46, 'bg'],
    ['rect', 0.00, 0.76, 1.00, 0.24, 'bg'],
    ['ellipse', 0.36, 0.30, 0.28, 0.28, 'fg'],
    ['round2SameRect', 0.28, 0.62, 0.44, 0.24, 'fg'],
  ],
  vlogger: [ // person framed by a phone held sideways on a grip
    ['roundRect', 0.00, 0.14, 1.00, 0.50, 'fg'],
    ['ellipse', 0.06, 0.31, 0.05, 0.16, 'bg'],
    ['ellipse', 0.89, 0.31, 0.05, 0.16, 'bg'],
    ['ellipse', 0.39, 0.25, 0.22, 0.22, 'bg'],
    ['ellipse', 0.42, 0.28, 0.16, 0.16, 'fg'],
    ['round2SameRect', 0.33, 0.44, 0.34, 0.20, 'bg'],
    ['round2SameRect', 0.36, 0.47, 0.28, 0.14, 'fg'],
    ['roundRect', 0.38, 0.60, 0.24, 0.22, 'bg'],
    ['roundRect', 0.41, 0.63, 0.18, 0.16, 'fg'],
    ['rect', 0.45, 0.86, 0.10, 0.09, 'fg'],
  ],
  likes: [ // thumbs-up knocked out of a filled disc
    ['ellipse', 0.00, 0.00, 1.00, 1.00, 'fg'],
    ['round2SameRect', 0.44, 0.20, 0.20, 0.28, 'bg'],
    ['roundRect', 0.32, 0.42, 0.38, 0.34, 'bg'],
    ['rect', 0.21, 0.44, 0.09, 0.32, 'bg'],
  ],
  followers: [ // a trio of figures knocked out of a filled disc
    ['ellipse', 0.00, 0.00, 1.00, 1.00, 'fg'],
    ['ellipse', 0.24, 0.22, 0.15, 0.15, 'bg'],
    ['round2SameRect', 0.18, 0.38, 0.27, 0.13, 'bg'],
    ['ellipse', 0.61, 0.22, 0.15, 0.15, 'bg'],
    ['round2SameRect', 0.55, 0.38, 0.27, 0.13, 'bg'],
    ['ellipse', 0.42, 0.50, 0.15, 0.15, 'bg'],
    ['round2SameRect', 0.36, 0.66, 0.27, 0.13, 'bg'],
  ],
};

function glyph(slide, kind, x, y, size, fg = WHITE, bg = INK) {
  GLYPHS[kind].forEach(([shape, gx, gy, gw, gh, role]) => {
    slide.addShape(shape, Object.assign(at(x + gx * size, y + gy * size, gw * size, gh * size), {
      fill: { color: role === 'bg' ? bg : fg },
      rectRadius: Math.min(gw, gh) * size * 0.3,
    }));
  });
}

/* ------------------------------------------------------------- slide bodies */

/** Full-bleed grey cover used for the opening and closing slides. */
function coverSlide(slide, heading) {
  slide.addShape('rect', Object.assign(at(0, 0, 13.333, 7.5), { fill: { color: INK, transparency: 50 } }));
  slide.addText(heading, Object.assign(at(0.808, 2.112, 11.448, 2.827), {
    fontFace: HEAD, fontSize: 162, charSpacing: 3, color: WHITE, align: 'center', valign: 'bottom',
  }));
  slide.addText(LOREM.aliqua, Object.assign(at(3.309, 6.156, 6.981, 0.787), {
    fontFace: TEXT, fontSize: 11, italic: true, color: WHITE, align: 'center', lineSpacingMultiple: 2, valign: 'top',
  }));
  [[12.256, -0.525], [-0.595, 6.353]].forEach(([x, y]) => {
    slide.addShape('ellipse', Object.assign(at(x, y, 1.673, 1.673), { fill: { color: INK }, line: { color: WHITE, width: 12 } }));
  });
}

const SLIDES = [];

// 1 - cover
SLIDES.push(s => coverSlide(s, 'CELEBRITY'));

// 2 - photo left, four descriptions right (bottom half reversed out on black)
SLIDES.push(s => {
  s.addShape('rect', Object.assign(at(0, 4.66, 13.333, 2.84), { fill: { color: INK } }));
  title(s, 'Subtitle Section Here', 3.044, 0.508, 7.245, 0.74, { align: 'center' });
  photo(s, 1.44, 1.47, 3.16, 6.03); // cut-out portrait, bleeding off the bottom edge
  [[5.6, 2.393, INK], [9.268, 2.393, INK], [5.6, 5.162, WHITE], [9.268, 5.162, WHITE]].forEach(([x, y, c]) => {
    label(s, 'Description Here', x, y, { color: c });
    para(s, LOREM.short, x + 0.111, y + 0.519, 2.765, 1.157, { color: c });
  });
});

// 3 - two pull quotes
SLIDES.push(s => {
  const quote = (text, author, y) => s.addText([
    { text, options: { breakLine: true } },
    { text: '', options: { breakLine: true } },
    { text: author },
  ], Object.assign(at(0.825, y, 5.425, 2.322), {
    fontFace: TEXT, fontSize: 18, italic: true, color: '101010', lineSpacingMultiple: 1.5, valign: 'top',
  }));
  quote('\u201CIn The Social Jungle Of Human Existence, There Is No Feeling Of Being Alive Without A Sense Of Identity.\u201D', '- Erik Erikson', 0.937);
  quote('\u201CWhat\u2019s the whole point of being pretty on the outside when you\u2019re so ugly on the inside?\u201D', '- Jess C. Scott', 4.083);
  badge(s, 10.659, 1.548);
});

// 4 - four personas in a row
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 3.044, 0.508, 7.245, 0.74, { align: 'center' });
  const copy = [LOREM.short, 'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut.'];
  [0.785, 3.865, 6.945, 10.025].forEach((x, i) => {
    label(s, 'Persona #' + (i + 1), x, 1.847, { align: 'center' });
    para(s, copy[i === 0 ? 0 : 1], x - 0.265, 2.35, 3.001, 1.157, { align: 'center' });
  });
});

// 5 - photo left, stacked copy right
SLIDES.push(s => {
  title(s, ['Subtitle', 'Section Here'], 7.999, 0.993, 4.144, 1.569);
  para(s, LOREM.consequat, 7.999, 2.786, 4.144, 1.898);
  para(s, LOREM.aliqua, 7.999, 4.909, 4.144, 1.157);
  badge(s, 5.113, 3.013);
});

// 6 - banner photo above, two columns of copy
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 3.044, 4.206, 7.245, 0.74, { align: 'center' });
  [0.583, 6.908].forEach(x => para(s, LOREM.consequat, x, 5.216, 5.843, 1.527));
});

// 7 - three stacked personas, heading beneath
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 0.866, 5.852, 4.53, 0.801);
  [0.831, 2.416, 3.972].forEach((y, i) => {
    label(s, 'Persona #' + (i + 1), 0.785, y);
    para(s, LOREM.dolore, 0.896, y + 0.503, 4.945, 0.787);
  });
});

// 8 - photo grid left, two pill-tagged paragraphs right
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 7.284, 1.225, 5.208, 0.801);
  [2.455, 4.636].forEach((y, i) => {
    pill(s, 'Persona #' + (i + 1), 7.429, y, 1.392, 0.513);
    para(s, LOREM.nostrud, 7.332, y + 0.626, 4.827, 1.157);
  });
});

// 9 - four personas over a portrait strip
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 3.044, 0.508, 7.245, 0.74, { align: 'center' });
  [[0.856, 1.714], [3.809, 1.732], [6.763, 1.696], [9.716, 1.714]].forEach(([x, y], i) => {
    label(s, 'Persona #' + (i + 1), x, y);
    para(s, LOREM.colore, x + 0.111, y + 0.503, 2.562, 1.527);
  });
});

// 10 - copy left with CTA, photo right
SLIDES.push(s => {
  title(s, ['Subtitle', 'Section Here'], 0.77, 0.721, 4.144, 1.569);
  label(s, 'Description Here', 0.679, 2.603);
  para(s, LOREM.commodo, 0.799, 3.109, 4.116, 1.898);
  pill(s, 'Learn More', 0.872, 5.497, 1.396, 0.513);
});

// 11 - photo left, 2x2 personas right
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 7.601, 5.937, 4.643, 0.74);
  [[7.521, 1.206], [10.248, 1.206], [7.521, 3.296], [10.248, 3.296]].forEach(([x, y], i) => {
    label(s, 'Persona #' + (i + 1), x, y);
    para(s, LOREM.tempor, x + 0.095, y + 0.503, 2.36, 1.157);
  });
});

// 12 - photo strip above, headline plus side note
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 0.903, 4.078, 6.084, 0.801);
  para(s, LOREM.consequat, 0.951, 5.008, 6.084, 1.527);
  label(s, 'Personas', 10.462, 4.151);
  para(s, LOREM.aliqua, 10.573, 4.653, 2.471, 1.898);
});

// 13 - two copy blocks with progress bars
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 7.418, 0.621, 4.979, 0.801);
  label(s, 'Description Here', 7.327, 1.764);
  para(s, LOREM.commodo, 7.447, 2.27, 4.979, 1.527);
  progress(s, 'Efficiency', 87, 7.566, 4.658, 4.809);
  label(s, 'Description Here', 7.327, 5.359);
  para(s, 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore. et dolore magna aliqua ut enim ad minim veniam quis nostrud.', 7.447, 5.865, 4.979, 1.157);
  progress(s, 'Personas', 75, 0.476, 6.904, 6.162);
});

// 14 - bulleted list
SLIDES.push(s => {
  title(s, ['Subtitle', 'Section Here'], 0.77, 0.564, 4.144, 1.569);
  [2.358, 3.74, 5.123].forEach(y => {
    para(s, LOREM.aliqua, 0.77, y, 4.659, 1.157, { bullet: { characterCode: '2022', indent: 13.5 } });
  });
});

// 15 - dark band with a soft left-to-right gradient
SLIDES.push(s => {
  // Band ramps from light grey on the left to solid black at mid-slide.
  const bandY = 2.727, bandH = 4.273, steps = 72, ramp = 6.667;
  for (let i = 0; i < steps; i++) {
    const v = Math.round(217 - 204 * i / (steps - 1)).toString(16).toUpperCase().padStart(2, '0');
    s.addShape('rect', Object.assign(at(i * ramp / steps, bandY, ramp / steps + 0.02, bandH), { fill: { color: v + v + v } }));
  }
  s.addShape('rect', Object.assign(at(ramp, bandY, 13.333 - ramp, bandH), { fill: { color: INK } }));
  title(s, 'Subtitle Section Here', 3.044, 0.508, 7.245, 0.74, { align: 'center' });
  para(s, LOREM.aliqua, 3.309, 1.325, 6.981, 0.787, { align: 'center', italic: true, color: INK });
  label(s, 'Description Here', 7.393, 3.285, { color: WHITE });
  para(s, LOREM.consequat, 7.513, 3.792, 5.09, 1.527, { color: WHITE });
  para(s, LOREM.aliqua, 7.513, 5.381, 5.09, 0.787, { color: WHITE });
});

// 16 - two offset copy blocks
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 5.469, 0.523, 4.771, 0.74);
  [[6.073, 1.732, 6.193, 2.238], [6.907, 4.241, 7.026, 4.747]].forEach(([lx, ly, px, py]) => {
    label(s, 'Description Here', lx, ly);
    para(s, LOREM.consequat, px, py, 5.601, 1.527);
  });
});

// 17 - three photograph notes, badge bottom right
SLIDES.push(s => {
  [0.815, 2.416, 3.988].forEach((y, i) => {
    label(s, 'Photograph #' + (i + 1), 0.722, y);
    para(s, LOREM.dolore, 0.833, y + 0.503, 4.818, 0.787);
  });
  title(s, 'Subtitle Section Here', 0.803, 5.852, 4.53, 0.801);
  badge(s, 10.763, 4.971);
});

// 18 - team grid
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 3.044, 0.51, 7.245, 0.74, { align: 'center' });
  const team = [['Chloe Joyzee', 0.831, 6.217, 1.751], ['Amanda Smith', 3.887, 6.265, 1.751],
    ['Ellen Moralen', 6.944, 6.265, 1.763], ['Ibiza Carlos', 9.992, 6.271, 1.751]];
  team.forEach(([name, x, ny, py]) => {
    para(s, LOREM.sedDo, x, py, 2.471, 1.157, { align: 'center' });
    label(s, name, x, ny, { align: 'center' });
  });
});

// 19 - two identity blocks beside stacked photos
SLIDES.push(s => {
  title(s, 'Subtitle Section Here', 7.418, 0.621, 4.979, 0.801);
  [[1.796, 2.302], [4.726, 5.233]].forEach(([ly, py]) => {
    label(s, 'Personal Identity', 7.327, ly);
    para(s, LOREM.commodo, 7.447, py, 4.979, 1.527);
  });
});

// 20 - copy left, six-up photo grid right
SLIDES.push(s => {
  title(s, ['Subtitle', 'Section Here'], 0.754, 0.564, 3.373, 1.569);
  para(s, LOREM.commodo, 0.754, 2.358, 3.373, 2.268);
  pill(s, 'Personal Identity', 0.822, 4.984);
});

// 21 - full-height photo left, 2x2 identity notes right
SLIDES.push(s => {
  pill(s, 'Personal Identity', 0.416, 6.619);
  title(s, 'Subtitle Section Here', 7.365, 0.646, 5.306, 0.74);
  [[7.301, 1.854, 1.898, LOREM.incididunt], [10.13, 1.854, 1.898, LOREM.incididunt],
    [7.301, 4.628, 1.527, LOREM.incididuntShort], [10.13, 4.628, 1.527, LOREM.incididuntShort]].forEach(([x, y, h, t]) => {
    label(s, 'Personal Identity', x, y);
    para(s, t, x + 0.079, y + 0.501, 2.36, h);
  });
});

// 22 - app promo with two phone mock-ups
SLIDES.push(s => {
  title(s, 'Find All Informations on App', 0.77, 0.721, 4.451, 1.569, { lineSpacingMultiple: 1.2 });
  label(s, 'Description Here', 0.679, 2.603);
  para(s, LOREM.commodo, 0.799, 3.108, 4.966, 1.527);

  const store = (x, caption, drawMark) => {
    s.addShape('roundRect', Object.assign(at(x, 5.486, 1.991, 0.76), { fill: { color: INK }, rectRadius: 0.1 }));
    drawMark(x + 0.229, 5.626);
    s.addText(caption, Object.assign(at(x + 0.68, 5.584, 1.311, 0.528), {
      fontFace: TEXT, fontSize: 10.5, color: WHITE, lineSpacingMultiple: 1.2, valign: 'middle',
    }));
  };
  store(0.799, 'Download on GooglePlay', (x, y) => {
    s.addShape('triangle', Object.assign(at(x, y, 0.33, 0.42), { fill: { color: WHITE }, rotate: 90 }));
  });
  store(3.23, 'Download on AppStore', (x, y) => {
    s.addShape('ellipse', Object.assign(at(x + 0.02, y + 0.06, 0.3, 0.3), { fill: { color: WHITE } }));
    s.addShape('ellipse', Object.assign(at(x + 0.16, y - 0.04, 0.12, 0.12), { fill: { color: WHITE } }));
  });

  [6.532, 9.732].forEach(x => { // phone mock-up: metal body, bezel, screen, notch
    s.addShape('roundRect', Object.assign(at(x, 1.254, 2.49, 4.992), { fill: { color: '8C8C8C' }, rectRadius: 0.14 }));
    s.addShape('roundRect', Object.assign(at(x + 0.05, 1.304, 2.39, 4.892), { fill: { color: '1A1A1A' }, rectRadius: 0.13 }));
    s.addShape('roundRect', Object.assign(at(x + 0.13, 1.384, 2.23, 4.732), { fill: { color: WHITE }, rectRadius: 0.1 }));
    s.addShape('roundRect', Object.assign(at(x + 0.79, 1.384, 0.91, 0.17), { fill: { color: '1A1A1A' }, rectRadius: 0.07 }));
  });
});

// 23 - desktop mock-up with two tagged paragraphs
SLIDES.push(s => {
  // Desktop mock-up: dark bezel, white screen, brushed foot and a soft shadow.
  s.addShape('rect', Object.assign(at(0.83, 0.72, 6.4, 4.28), { fill: { color: '141414' } }));
  s.addShape('rect', Object.assign(at(1.027, 1.022, 5.975, 3.691), { fill: { color: WHITE } }));
  s.addShape('rect', Object.assign(at(0.83, 5.0, 6.4, 0.55), { fill: { color: 'AFAFAF' } }));
  s.addShape('trapezoid', Object.assign(at(3.2, 5.55, 1.9, 0.8), { fill: { color: 'CBCBCB' }, flipV: true }));
  s.addShape('ellipse', Object.assign(at(2.55, 6.2, 3.2, 0.42), { fill: { color: 'DDDDDD' } }));

  title(s, 'Browse Via Web', 7.767, 0.79, 4.451, 0.801, { lineSpacingMultiple: 1.2 });
  [[2.096, 2.742], [4.317, 4.967]].forEach(([py, ty]) => {
    pill(s, 'Personal Identity', 7.901, py);
    para(s, LOREM.aliqua, 7.799, ty, 4.568, 1.157);
  });
});

// 24 - four icon discs beside descriptions
SLIDES.push(s => {
  title(s, 'Infographic Here', 3.044, 0.527, 7.245, 0.74, { align: 'center' });
  const discs = [[1.053, 2.153, INK, 'influencer', 1.372, 2.44], [3.39, 2.153, MID, 'vlogger', 3.67, 2.569],
    [1.014, 4.517, MID, 'likes', 1.333, 4.837], [3.35, 4.517, INK, 'followers', 3.67, 4.837]];
  discs.forEach(([x, y, c, kind, gx, gy]) => {
    s.addShape('ellipse', Object.assign(at(x, y, 1.917, 1.917), { fill: { color: c } }));
    glyph(s, kind, gx, gy, 1.278, WHITE, c);
  });
  const copy = [['Influencer', 5.858, 2.05, 5.966, 2.551], ['Vlogger', 9.158, 2.041, 9.294, 2.542],
    ['Likes', 5.858, 4.344, 5.966, 4.845], ['Followers', 9.186, 4.336, 9.294, 4.837]];
  copy.forEach(([name, lx, ly, px, py]) => {
    label(s, name, lx, ly);
    para(s, LOREM.incididunt, px, py, 2.943, 1.527);
  });
});

// 25 - four banner blocks with captions beneath
SLIDES.push(s => {
  const blocks = [[0, INK, 'influencer', 0.74, 2.231, 0.182, 4.397, 0.298, 4.902, 2.668],
    [3.453, MID, 'vlogger', 4.193, 2.262, 3.237, 4.388, 3.353, 4.893, 3.066],
    [6.917, INK, 'likes', 7.657, 2.306, 6.692, 4.397, 6.808, 4.902, 3.066],
    [10.383, MID, 'followers', 11.123, 2.306, 10.161, 4.388, 10.277, 4.893, 2.668]];
  const copy = ['Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do. eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim.',
    'Lorem ipsum dolor sit amet consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniamiso quis.'];
  blocks.forEach(([x, c, kind, gx, gy, lx, ly, px, py, pw], i) => {
    s.addShape('rect', Object.assign(at(x, 1.987, 2.95, 2.108), { fill: { color: c } }));
    glyph(s, kind, gx, gy, 1.471, WHITE, c);
    label(s, 'Description Here', lx, ly);
    para(s, copy[i === 0 || i === 3 ? 0 : 1], px, py, pw, 1.898);
  });
  title(s, 'Infographic Here', 3.044, 0.527, 7.245, 0.74, { align: 'center' });
});

// 26 - zig-zag timeline of icon tiles
SLIDES.push(s => {
  title(s, 'Infographic Here', 3.044, 0.527, 7.245, 0.74, { align: 'center' });
  const steps = [
    { tile: [2.271, 5.349], color: INK, kind: 'influencer', dot: [2.655, 4.367], line: [2.791, 4.758], text: [1.364, 2.841], lbl: [1.549, 2.336] },
    { tile: [4.887, 2.35], color: MID, kind: 'vlogger', dot: [5.273, 4.106], line: [5.406, 3.523], text: [3.94, 5.203], lbl: [4.126, 4.698] },
    { tile: [7.459, 5.339], color: INK, kind: 'likes', dot: [7.843, 4.357], line: [7.978, 4.748], text: [6.552, 2.83], lbl: [6.737, 2.326] },
    { tile: [10.075, 2.34], color: MID, kind: 'followers', dot: [10.46, 4.096], line: [10.594, 3.513], text: [9.128, 5.193], lbl: [9.313, 4.688] },
  ];
  steps.forEach(st => {
    s.addShape('rect', Object.assign(at(st.tile[0], st.tile[1], 1.039, 1.039), { fill: { color: st.color } }));
    glyph(s, st.kind, st.tile[0] + 0.19, st.tile[1] + 0.19, 0.658, WHITE, st.color);
    s.addShape('rect', Object.assign(at(st.dot[0], st.dot[1], 0.271, 0.271), { fill: { color: st.color } }));
    s.addShape('line', Object.assign(at(st.line[0], st.line[1], 0, 0.471), { line: { color: st.color, width: 2.25 } }));
    label(s, 'Description Here', st.lbl[0], st.lbl[1], { align: 'center' });
    para(s, LOREM.short, st.text[0], st.text[1], 2.841, 1.157, { align: 'center' });
  });
});

// 27 - single hero icon with four surrounding notes
SLIDES.push(s => {
  s.addShape('ellipse', Object.assign(at(4.864, 2.683, 3.605, 3.605), { fill: { color: INK } }));
  glyph(s, 'influencer', 5.471, 3.258, 2.391, WHITE, INK);
  title(s, 'Infographic Here', 3.044, 0.527, 7.245, 0.74, { align: 'center' });
  [[0.882, 2.417, 0.998, 2.922], [0.882, 4.626, 0.998, 5.131],
    [8.932, 2.543, 9.048, 3.048], [8.932, 4.752, 9.048, 5.257]].forEach(([lx, ly, px, py]) => {
    label(s, 'Description Here', lx, ly);
    para(s, LOREM.labore, px, py, 3.391, 1.157);
  });
});

// 28 - four tall cards
SLIDES.push(s => {
  const cards = [[0.986, 2.249, INK, 'influencer', 1.767, 2.49, 'Influencer', 0.987, 3.614, 1.113, 4.151],
    [3.949, 2.249, MID, 'vlogger', 4.73, 2.49, 'Vlogger', 3.95, 3.614, 4.076, 4.151],
    [6.913, 2.249, INK, 'likes', 7.689, 2.476, 'Likes', 6.913, 3.61, 7.039, 4.147],
    [9.876, 2.248, MID, 'followers', 10.656, 2.476, 'Followers', 9.877, 3.61, 10.003, 4.147]];
  cards.forEach(([x, y, c, kind, gx, gy, name, lx, ly, px, py]) => {
    s.addShape('rect', Object.assign(at(x, y, 2.471, 4.243), { fill: { color: c } }));
    glyph(s, kind, gx, gy, 0.91, WHITE, c);
    label(s, name, lx, ly, { align: 'center', color: WHITE });
    para(s, LOREM.labore, px, py, 2.22, 1.898, { align: 'center', color: WHITE });
  });
  title(s, 'Infographic Here', 3.044, 0.527, 7.245, 0.74, { align: 'center' });
});

// 29 - stat callouts around a 2x2 icon cluster
SLIDES.push(s => {
  [['influencer', 4.685, 2.315, INK], ['vlogger', 6.894, 2.41, MID],
    ['likes', 4.671, 4.509, MID], ['followers', 6.894, 4.523, INK]]
    .forEach(([kind, x, y, c]) => glyph(s, kind, x, y, 1.754, c, WHITE));
  const stats = [['$45.000 / month', 0.801, 2.396, 0.917, 2.901], ['30k likes', 0.787, 4.494, 0.903, 4.999],
    ['Daily Vlogger', 9.071, 2.354, 9.187, 2.859], ['702k followers', 9.071, 4.508, 9.187, 5.013]];
  stats.forEach(([name, lx, ly, px, py]) => {
    label(s, name, lx, ly);
    para(s, LOREM.labore, px, py, 3.391, 1.157);
  });
  title(s, 'Infographic Here', 3.044, 0.527, 7.245, 0.74, { align: 'center' });
});

// 30 - closing cover
SLIDES.push(s => coverSlide(s, 'THANK YOU'));

// 31 / 32 - map libraries. Each original silhouette is hundreds of vector
// paths; here every landmass is stood in for by one soft grey blob occupying
// its original footprint.
function mapSheet(regions) {
  return s => regions.forEach(([x, y, w, h]) => {
    s.addShape('cloud', Object.assign(at(x, y, w, h), { fill: { color: RULE }, line: { color: RULE, width: 1 } }));
  });
}
SLIDES.push(mapSheet([[1.458, 1.049, 2.216, 2.492], [4.564, 1.193, 3.502, 2.205], [9.181, 1.224, 3.028, 2.143],
  [1.568, 3.891, 1.997, 2.789], [4.936, 4.222, 2.758, 2.127], [9.338, 4.098, 2.712, 2.374]]));
SLIDES.push(mapSheet([[0.964, 1.394, 3.088, 2.057], [4.452, 1.116, 1.914, 2.614], [6.864, 0.847, 2.172, 3.152],
  [9.521, 1.283, 2.972, 2.28], [0.928, 4.501, 3.16, 1.658], [4.513, 4.304, 1.793, 2.051],
  [6.722, 3.979, 2.455, 2.7], [9.438, 4.098, 3.139, 2.462]]));

/* --------------------------------------------------------------------- main */

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'DECK', width: 13.333, height: 7.5 });
pptx.layout = 'DECK';
pptx.author = 'pptxgenjs';
pptx.title = 'Celebrity';

SLIDES.forEach(build => {
  const slide = pptx.addSlide();
  slide.background = { color: WHITE };
  build(slide);
});

pptx.writeFile({ fileName: path.join(__dirname, '09dfca34-01d6-413f-a747-40cc90b7a7a2_grok_final.pptx') })
  .then(f => console.log('wrote ' + f));
