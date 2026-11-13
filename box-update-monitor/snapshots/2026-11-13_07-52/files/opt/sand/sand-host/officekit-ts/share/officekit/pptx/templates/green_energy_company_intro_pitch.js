/**
 * Green Energy - 30-slide presentation template, rebuilt with PptxGenJS.
 *
 * Palette: dark charcoal panels, cream page ground, lime accents.
 * The recurring 'circuit' artwork (dots + connecting wires) is described as
 * small ASCII maps -- 'o' hollow dot, '*' filled dot, '-' / '|' wires -- laid
 * out on a 0.09764" lattice, so every motif stays readable in source.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- theme
const DARK  = '232625';   // charcoal panels / body text
const LIME  = 'B3D929';   // accent green
const CREAM = 'D2D9B8';   // page background

const BODY  = 'Poppins';
const SEMI  = 'Poppins SemiBold';
const BLACK = 'Poppins Black';
const XBOLD = 'Poppins ExtraBold';

const CELL = 0.09764;     // artwork lattice pitch (inches)
const DOT  = 0.117;       // circuit dot diameter

// ------------------------------------------------------- circuit artwork
// Each entry is an ASCII map of one connected motif. Columns/rows are CELL
// apart; 'o' = outlined dot, '*' = filled dot, '-'/'|' = 1pt wires.
const ART = [
  /*  0 */ [
    '               *',
    '               |',
    '               |',
    '               |',
    '               |',
    '               |',
    '               |',
    '               |',
    '',
    '',
    '       o------ *',
    '',
    '               *',
    '',
    '       o------ o',
    '       |',
    '       |',
    '       |',
    '       |',
    '       |',
    '',
    '       *------ o * * *',
    '',
    '------ o       o * * *',
    '       |',
    '       |       * o * * ----- o',
    '       |',
    '       |       * * * o       |',
    '                             |',
    '       *       o * * *       |',
    '                             |',
    '                             |',
    '                             |',
    '                             |',
    '                             |',
    '                             |',
    '                             |',
    '',
    '                             *',
  ],
  /*  1 */ [
    '         o',
    '         |',
    '         |',
    '         |',
    '',
    '',
    'o ------ *',
  ],
  /*  2 */ [
    'o *--------  o',
    '|',
    '|',
    '|',
    '|',
    '|',
    '',
    'o',
    '|',
    '|',
    '|',
    '|',
    '',
    '*',
  ],
  /*  3 */ [
    '* ------ o  o *',
  ],
  /*  4 */ [
    'o ------ *',
  ],
  /*  5 */ [
    '* ----- o------ o',
    '',
    '        |',
    '        |',
    '        |',
    '        |',
    '        |',
    '',
    '        *------ o * * *',
    '',
    '                o * * *',
    '',
    '                * o * *',
    '',
    '                * * * o',
    '',
    '                o * * *',
  ],
  /*  6 */ [
    '* -------  * o ------ *',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '',
    '',
    '             o ------ *',
  ],
  /*  7 */ [
    '* o * * o ------- o *',
  ],
  /*  8 */ [
    '               *',
    '',
    '               |',
    '               |',
    'o ------ *     |',
    '         |     |',
    '         |     |',
    '         |     |',
    '               |',
    '',
    '         * o * * o -------  * o ------ * * --------- o o  *',
  ],
  /*  9 */ [
    'o------ *',
    '',
    '        *',
    '',
    'o------ o',
  ],
  /* 10 */ [
    '*---------  o o *',
  ],
  /* 11 */ [
    'o * * *',
    '',
    'o * * *',
    '',
    '* o * *',
    '',
    '* * * o',
    '',
    'o * * *',
  ],
  /* 12 */ [
    'o------  *',
  ],
  /* 13 */ [
    '*--------  * o------  *',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '                      |',
    '',
    '',
    '             o ------ *',
  ],
  /* 14 */ [
    'o------- *',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '',
    '',
    '*',
  ],
  /* 15 */ [
    '         *',
    '         |',
    '         |',
    '         |',
    '         |',
    '         |',
    '         |',
    '         |',
    '',
    '',
    'o * ---- o',
  ],
  /* 16 */ [
    'o------- *',
    '',
    '         o',
    '',
    'o------- o',
    '|        |',
    '|        |',
    '|        |',
    '|        |',
    '|        |',
    '         |',
    '*        |',
    '         |',
    '         |',
    '',
    '         *',
  ],
  /* 17 */ [
    'o------- *',
    '',
    '         *',
    '',
    'o------- o',
  ],
  /* 18 */ [
    '            o',
    '            |',
    '            |',
    '            |',
    'o * * *     |',
    '            |',
    'o * * *',
    '',
    '* o * *---  *',
    '',
    '* * * o     o',
    '            |',
    'o * * *     |',
    '            |',
    '            |',
    '',
    '   * ------ *',
  ],
  /* 19 */ [
    '         *',
    '',
    ' o------ o *',
    '         |',
    '         |',
    '         |',
    '         |',
    '',
    '*------',
    '         o',
    '|        |',
    '|        |',
    '|        |',
    '|        |',
    '|',
    '|        *',
    '|',
    '',
    '',
    '*------  o',
    '',
    '*',
    'o',
  ],
  /* 20 */ [
    'o',
    '',
    '*------  o',
    '         |',
    '         |',
    '',
    '',
    '* o----  o',
  ],
  /* 21 */ [
    '* ----- o',
    '',
    '        |',
    '        |',
    '        |',
    '        |',
    '        |',
    '',
    '        *',
  ],
  /* 22 */ [
    '*------ o * * *',
    '',
    '        o * * *',
    '',
    '        * o * *',
    '',
    '        * * * o',
    '',
    '        o * * *',
  ],
  /* 23 */ [
    '*',
    '',
    '|',
    '|',
    '|',
    '|',
    '|',
    '',
    'o ------ *',
  ],
  /* 24 */ [
    '*',
    '',
    '|',
    '|',
    '|',
    '|',
    '',
    '*',
  ],
  /* 25 */ [
    '         o *------ *',
    '         |         |',
    '         |         |',
    '         |         |',
    '         |         |',
    '         |         |',
    '         |         |',
    '         |',
    '         |         o',
    '         |',
    '',
    '',
    'o ------ *',
  ],
  /* 26 */ [
    '         * ------ o',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '',
    '',
    '*---------  o o * *',
  ],
  /* 27 */ [
    '* o ------ * * --------- o o  *',
  ],
  /* 28 */ [
    'o------ *',
    '',
    '        *',
    '',
    'o------ o',
    '|',
    '|',
    '',
    '*',
  ],
  /* 29 */ [
    '|',
    '|       o------ *',
    '|',
    '|               *',
    '|',
    '|       o------ o',
    '|       |',
    '        |',
    '        *',
    'o------',
  ],
  /* 30 */ [
    '---------  o o *',
  ],
  /* 31 */ [
    '*---------  o o *             * --------- o o *',
    '',
    '                *------ o * * *',
    '',
    '                        o * * *',
    '',
    '                        * o * *',
    '',
    '                        * * * o',
    '',
    '                        o * * *',
  ],
  /* 32 */ [
    '       o------ *',
    '',
    '               *',
    '',
    '       o------ o',
    '       |',
    '       |',
    '',
    '------ *',
  ],
  /* 33 */ [
    '        o   o * * *',
    '        |',
    '        |   o * * *',
    '        |',
    '        |   * o * *',
    '',
    '        *   * * * o',
    '',
    'o------ o * o * * * -----------',
  ],
  /* 34 */ [
    '*----- o',
    '|',
    '|',
    '|',
    '',
    'o',
  ],
  /* 35 */ [
    '       *',
    '',
    '       *',
    '',
    '------ o',
  ],
  /* 36 */ [
    '         *',
    '',
    'o------  o *',
    '         |',
    '         |',
    '         |',
    '         |',
    '',
    '*------',
    '         o',
    '|        |',
    '|        |',
    '|        |',
    '|        |',
    '|',
    '|        *',
    '|',
    '',
    '',
    '*------  o',
    '',
    '*',
  ],
  /* 37 */ [
    '          *',
    '',
    'o ------  o *',
    '            |',
    '            |',
    '            |',
    '            |',
    '            |',
    '            |',
    'o * * *     |',
    '            |',
    'o * * *',
    '            *',
    '* o * *---',
    '',
    '* * * o     o',
    '            |',
    'o * * *     |',
    '            |',
    '            |',
    '     ------ *',
    '   *',
  ],
  /* 38 */ [
    'o o *             * --------- o o *',
    '',
    '    *------ o * * *',
    '',
    '            o * * *',
    '',
    '            * o * *',
    '',
    '            * * * o',
    '',
    '            o * * *',
  ],
  /* 39 */ [
    '         o------ *',
    '',
    'o------  *       *',
    '',
    '*        o------ o       o',
    '         |               |',
    '         |               |',
    '',
    '         *---------  o o *',
    '',
    '   o * * *',
    '',
    '   o * * *',
    '',
    '   * o * *',
    '',
    '   * * * o',
    '',
    '   o * * *',
  ],
  /* 40 */ [
    'o----  * * --------- o o *',
  ],
  /* 41 */ [
    '        *',
    '',
    'o ----  o *',
    '|       |',
    '|       |',
    '|       |',
    '|       |',
    '|       |',
    '',
    '',
    '* ----  *',
    '',
    '        o',
    '        |',
    '        |',
    '        |',
    '        |',
    '        |',
    '        |',
    '',
    '       *',
    ' -----',
  ],
  /* 42 */ [
    'o---------- *',
    '            |',
    '*----- o    |     o---- * o',
    '            |     |',
    'o                 |',
    '',
    '     o  *-- o --- *---- *',
  ],
  /* 43 */ [
    '         *',
    '',
    '  o ---- o  *',
    '  |      |',
    '  |      |',
    '  |      |',
    '  |      |',
    '         |',
    'o * * *',
    '',
    'o * * *   *',
    '',
    '* o * *  o',
    '         |',
    '* * * o  |',
    '         |',
    'o * * *  |',
    '         |',
    '         |',
    '',
    '         *',
    '  *-----',
  ],
  /* 44 */ [
    '     *---  o---- o',
    '     -     |',
    '           |',
    '',
    '*--  o---  *',
  ],
  /* 45 */ [
    'o----- *',
    '',
    '*----- o',
    '',
    'o',
  ],
  /* 46 */ [
    'o *',
  ],
  /* 47 */ [
    '         o *    o ------ o',
    '           |    |',
    '           |    |',
    '',
    '---------  o--- *',
  ],
  /* 48 */ [
    '  *',
    '  |',
    '  |',
    '  |',
    '* |',
    '  |',
    'o |',
    '| |',
    '| |',
    '| |',
    '| |',
    '| |',
    '| |',
    '| |',
    '',
    '',
    '* o',
    '',
    'o',
    '',
    '*',
    '|',
    '|',
    '|',
    '|',
  ],
  /* 49 */ [
    'o-- *',
  ],
  /* 50 */ [
    '         *',
    '',
    '  o ---- o',
    '  |',
    '  |',
    '  |',
    '  |',
    '',
    '  o',
    '',
    '  *-  *',
    '',
    '  o   o',
    '  |',
    '  |',
    '  |',
    '',
    '  o----- *',
    '  |      |',
    '  |      |',
    '  |      |',
    '  |      |',
    '',
    'o * * *  o',
    '         |',
    'o * * *  |',
    '',
    '* o * *  o ----- *  o------ o',
    '                    |',
    '* * * o  *          |',
    '',
    'o * * *  * -- o --- * o *',
  ],
  /* 51 */ [
    '          *',
    '',
    'o ------  o *',
    '            |',
    '            |',
    '            |',
    '            |',
    '            |',
    '            |',
    'o * * *     |',
    '            |',
    'o * * *',
    '            *',
    '* o * *---',
    '',
    '* * * o     o',
    '            |',
    'o * * *     |',
    '            |',
    '            |',
    '            *',
    '  *-------',
  ],
  /* 52 */ [
    '  o',
    '',
    'o *-------  o',
    '            |    o------ o',
    '            |    |',
    '                 |',
    '',
    '  * o-----  o--- *',
  ],
  /* 53 */ [
    '                          *',
    '',
    '                          o',
    '',
    '                  o------ *',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '',
    '',
    '                  *---  *',
    '',
    '         *              o *',
    '                          |',
    'o------  o *              |',
    '                          |',
    '         |                |',
    '         |                |',
    '         |                |',
    '         |                |',
    '',
    '*------  o--------------  *',
    '|        |',
    '|        |                *',
    '|        |',
    '|        |                *',
    '|',
    '|        *------  *       o',
    '|',
    '                  |       |',
    '                  |       |',
    '*                 |       |',
    '',
    '*------  o      * o       *',
  ],
  /* 54 */ [
    '          *       o* *',
    '                  |',
    'o ------  o *     |',
    '',
    '            o---  *  o------ *',
    '            |        |',
    '            |        |',
    '            |        |',
    '            |        |',
    'o * * *     |        |',
    '            |        |',
    'o * * *              |',
    '            *        |',
    '* o * *---',
    '',
    '* * * o     o        *---  *',
    '            |        |',
    'o * * *     |        |     o *',
    '            |        |       |',
    '            |                |',
    '            *        o       |',
    '  *-------   -------         |',
    '                             |',
    '                             |',
    '                             |',
    '',
    '                             *',
    '',
    '                             *',
    '',
    '                             *',
    '',
    '                             o',
    '',
    '                     |       |',
    '                     |       |',
    '                     |       |',
    '',
    '                   * o ----- *',
  ],
  /* 55 */ [
    '  o',
    '',
    'o *-------  o',
    '            |',
    '            |',
    '',
    '',
    '* * o-----  o',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '',
    '*',
    '',
    '*',
    '',
    '*',
    '',
    'o * o',
    '',
    '|',
    '|',
    '|',
    '',
    '* o *',
    '',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '',
    'o   *',
    '',
    '    *',
    '',
    '    o------ *',
  ],
  /* 56 */ [
    '* ----- o',
    '',
    '        *',
    '',
    '        *',
  ],
  /* 57 */ [
    '         *',
    '',
    'o ------ o  *',
    '|        |',
    '|        |',
    '|        |',
    '|        |',
    '|        |',
    '|        |',
    '|        |',
    '|        |',
    '|',
    '| *      *',
    '|  -----',
    '|',
    '|        o',
    '|        |',
    '|        |',
    '|        |',
    '',
    'o        *  o',
    '  *-----',
  ],
  /* 58 */ [
    '  o',
    '',
    'o *-------  *',
    '|           |    o ------ o',
    '|           |    |',
    '|                |',
    '|',
    '|   o-----  *--- *',
    '|',
    '',
    '*',
  ],
  /* 59 */ [
    '                          *',
    '',
    '                          o',
    '',
    '                  o------ *',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '                  |',
    '',
    '',
    '                  *---  *',
    '',
    '         *              o *',
    '                          |',
    ' ------  o *              |',
    '                          |',
    '         |                |',
    '         |                |',
    '         |                |',
    '         |                |',
    '',
    '*------  o--------------  *',
    '|        |',
    '|        |                *',
    '|        |',
    '|        |    o * * *     *',
    '|',
    '|        *--- o * * *     o',
    '|',
    '              * o * *     |',
    '                          |',
    '*             * * * o     |',
    '',
    '*------  o    o * * *     *',
  ],
  /* 60 */ [
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '|',
    '',
    'o',
  ],
  /* 61 */ [
    '          *       o--*',
    '                  |',
    'o ------  o *     |',
    '',
    '            o---  *  o------ *',
    '            |        |',
    '            |        |',
    '            |        |',
    '            |        |',
    'o * * *     |        |',
    '            |        |',
    'o * * *              |',
    '            *        |',
    '* o * *---',
    '',
    '* * * o     o        *---  *',
    '            |        |',
    'o * * *     |        |     o *',
    '            |        |       |',
    '            |                |',
    '            *        o       |',
    '  *-------   -------         |',
    '                             |',
    '                             |',
    '                             |',
    '',
    '                             *',
    '',
    '                             *',
    '',
    '                             *',
    '',
    '                             o',
    '',
    '                     |       |',
    '                     |       |',
    '                     |       |',
    '',
    '                   * o ----- *',
  ],
  /* 62 */ [
    '                o----- *       o -------- *',
    '',
    '                |      |       |',
    '                |      |       |',
    '                |      |       |',
    '                |      |       |',
    '',
    '*---------  o o *      * o * * o -------- * o',
  ],
  /* 63 */ [
    '                         o',
    '                         |',
    '                         |',
    '                         |',
    '*---------  *            |',
    '            |',
    '            |        *   *',
    '',
    '            o------  o * o',
  ],
  /* 64 */ [
    '*---------  *',
  ],
  /* 65 */ [
    'o',
    '|',
    '|',
    '|',
    '|',
    '',
    '*',
    '',
    'o *',
  ],
  /* 66 */ [
    'o o *',
  ],
  /* 67 */ [
    '* o *',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '',
    '',
    'o   *',
    '',
    '    *',
    '',
    '    o',
  ],
  /* 68 */ [
    '|',
    '|',
    '|',
    '  *',
    '',
    '* o',
  ],
  /* 69 */ [
    '                                    * o *',
    '                                    |   |',
    '                                    |   |',
    '                                    |   |',
    '                                    |   |',
    '                                    |   |',
    '                                    |   |',
    '                                    |   |',
    '',
    '                                    o   *',
    '*---------  * *---------  o o *',
    '                            |       *',
    '                            |',
    '                            o------ o * o',
  ],
  /* 70 */ [
    '* o *',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '|   |',
    '',
    'o   *',
    '|',
    '|   *',
    '|',
    '|   o',
    '|',
    '',
    '',
    'o-----  *',
  ],
  /* 71 */ [
    '* o * * o--------- *  o',
  ],
  /* 72 */ [
    'o---------- *',
    '|',
    '|         * o',
    '|',
    '| * ----- o',
    '|',
    '|',
    '|',
    '| o ----- *',
    '| |',
    '| |',
    '| |',
    '| |',
    '| |',
    '| |',
    '',
    '',
    'o *',
  ],
  /* 73 */ [
    '  o------- *',
    '',
    'o *     o  o',
    '|',
    '|       *',
    '',
    'o------ o * o',
  ],
  /* 74 */ [
    '* o',
  ],
  /* 75 */ [
    '* o * * o ------',
  ],
  /* 76 */ [
    '* o * * o',
  ],
];

/** Paint one ASCII art motif with its top-left dot centred at (x, y). */
function art(s, id, x, y, color) {
  const rows = ART[id];
  const seen = new Set();
  const at = (c, r) => (rows[r] && c >= 0 && c < rows[r].length ? rows[r][c] : ' ');
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const ch = rows[r][c];
      if (ch === 'o' || ch === '*') {
        s.addShape('ellipse', {
          x: x + c * CELL - DOT / 2, y: y + r * CELL - DOT / 2, w: DOT, h: DOT,
          fill: ch === '*' ? { color: LIME } : { type: 'none' },
          line: { color: color, width: 1 },
        });
      } else if (ch === '-' && at(c - 1, r) !== '-') {
        let len = 1;
        while (at(c + len, r) === '-') len++;
        s.addShape('line', {
          x: x + c * CELL, y: y + r * CELL, w: len * CELL, h: 0,
          line: { color: color, width: 1 },
        });
      } else if (ch === '|' && at(c, r - 1) !== '|') {
        let len = 1;
        while (at(c, r + len) === '|') len++;
        s.addShape('line', {
          x: x + c * CELL, y: y + r * CELL, w: 0, h: len * CELL,
          line: { color: color, width: 1 },
        });
      }
    }
  }
}

// ------------------------------------------------------------- helpers
/** Solid charcoal panel. */
function panel(s, x, y, w, h, color) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: color || DARK } });
}

/** Thin horizontal rule. */
function rule(s, x, y, w, color) {
  s.addShape('line', { x: x, y: y, w: w, h: 0, line: { color: color, width: 1 } });
}

// ----------------------------------------------------------- typography
const MARGIN = [7.2, 7.2, 3.6, 3.6];   // [l, r, b, t] inset in points (0.1" / 0.05")

// Named run styles reused across the deck.
const STYLE = {
  title:      { fontFace: BLACK, fontSize: 44, color: DARK },
  hero:       { fontFace: BLACK, fontSize: 72, color: LIME },
  letter:     { fontFace: XBOLD, fontSize: 150, color: LIME, lineSpacingMultiple: 1 },
  pct:        { fontFace: 'Roboto Slab Black', fontSize: 44, color: LIME },
  stat:       { fontFace: SEMI, fontSize: 32, color: DARK },
  label:      { fontFace: SEMI, fontSize: 18, color: DARK },
  labelLight: { fontFace: SEMI, fontSize: 18, color: CREAM },
  body:       { fontFace: BODY, fontSize: 11, color: DARK, lineSpacingMultiple: 1.5 },
  bodyLight:  { fontFace: BODY, fontSize: 11, color: CREAM, lineSpacingMultiple: 1.5 },
  small:      { fontFace: BODY, fontSize: 9, color: DARK, lineSpacingMultiple: 1.5 },
};

/** Consecutive paragraphs sharing one style. */
const paras = list => list.map(t => ({ text: t, options: { breakLine: true } }));

/** Bulleted list sharing one style. */
const bullets = list => list.map(t => ({ text: t, options: { bullet: { indent: 13.5 } } }));

/**
 * Place a text box. `at` is [x, y, w, h]; `style` names a STYLE preset;
 * `body` is a plain string, or the output of paras()/bullets().
 */
function tx(s, at, style, body, more) {
  s.addText(body, Object.assign(
    { x: at[0], y: at[1], w: at[2], h: at[3], valign: 'top', margin: MARGIN },
    STYLE[style], more || {}));
}

/** Small right-aligned page marker in the top-right corner. */
function pageTag(s, label, color) {
  s.addText(label, {
    x: 12.372, y: 0.233, w: 0.763, h: 0.236, align: 'right', valign: 'top',
    fontFace: SEMI, fontSize: 8, color: color, margin: [5, 7, 5, 7],
  });
}

// ------------------------------------------------------------- layouts
// Dark background panels contributed by each slide layout.
const PANELS = {
  '21_Custom Layout': [[0, 4.019, 13.333, 3.481]],
  'Title and Content': [[11.317, 0, 2.016, 7.5]],
  'Section Header': [[0, 0, 2.016, 7.5]],
  'Two Content': [[0, 5.481, 13.333, 2.019]],
  'Comparison': [[0.926, 5.481, 4.686, 2.019]],
  '22_Custom Layout': [[0, 4.019, 13.333, 3.481]],
  'Blank': [[0, 5.152, 5.613, 2.348]],
  'Content with Caption': [[0, 0, 2.016, 7.5]],
  '2_Custom Layout': [[6.9, 5.429, 5.481, 2.071]],
  'Picture with Caption': [[1.583, 6.104, 5.335, 1.396]],
  'Title and Vertical Text': [[0, 0, 2.016, 4.044], [11.317, 3.456, 2.016, 4.044]],
  '4_Custom Layout': [[8.421, 5.307, 4.912, 2.193]],
  '5_Custom Layout': [[9.773, 4.665, 3.56, 2.835]],
  '23_Custom Layout': [[0, 4.019, 13.333, 3.481]],
  '24_Custom Layout': [[0, 0, 3.202, 3.75], [10.132, 3.75, 3.202, 3.75], [3.377, 3.75, 3.202, 3.75], [6.754, 0, 3.202, 3.75]],
  '25_Custom Layout': [[0, 0, 3.976, 7.5]],
  '26_Custom Layout': [[9.358, -0.033, 3.976, 7.533]],
  '27_Custom Layout': [[0, 0, 3.976, 7.5]],
  '28_Custom Layout': [[9.358, 0, 3.976, 7.5]],
  '29_Custom Layout': [[0, 6.034, 13.333, 1.466]],
  '30_Custom Layout': [[0, 4.019, 13.333, 3.481]],
  '31_Custom Layout': [[7.014, 5.307, 5.297, 2.193]],
  '32_Custom Layout': [[0, 0, 5.613, 1.953]],
  '33_Custom Layout': [[0, 6.034, 13.333, 1.466]],
  '34_Custom Layout': [[0, 4.019, 13.333, 3.481]],
  '18_Custom Layout': [[7.014, 0, 6.319, 2.478]],
  '19_Custom Layout': [[11.387, 2.163, 1.946, 4.003]],
  '20_Custom Layout': [[0, 2.311, 13.333, 1.382]],
};

// The 45-degree rotated dot rosette that recurs across the deck.
// Positions are the rosette centre; the whole cluster sits at ~315 degrees.
const ROSETTE = ['*o*', '***', '*o*'];
const ROSETTE_PITCH = 0.195;

/** Paint the rotated dot rosette centred on (cx, cy). */
function rosette(s, cx, cy) {
  const rad = (315.3 * Math.PI) / 180;
  const cos = Math.cos(rad), sin = Math.sin(rad);
  ROSETTE.forEach((row, r) => {
    for (let c = 0; c < row.length; c++) {
      const ux = (c - 1) * ROSETTE_PITCH, uy = (r - 1) * ROSETTE_PITCH;
      s.addShape('ellipse', {
        x: cx + ux * cos - uy * sin - DOT / 2,
        y: cy + ux * sin + uy * cos - DOT / 2,
        w: DOT, h: DOT,
        fill: row[c] === '*' ? { color: LIME } : { type: 'none' },
        line: { color: LIME, width: 1 },
      });
    }
  });
}

// Rosette centres contributed by each layout.
const ROSETTES = {
  'Title Slide': [[12.178,1.077], [9.772,3.643]],
  'Title and Content': [[12.172,1.818]],
  'Section Header': [[0.542,1.403]],
  'Two Content': [[4.805,6.506]],
  'Comparison': [[4.874,6.243]],
  'Blank': [[4.874,6.243], [0.637,5.847]],
  'Content with Caption': [[0.542,1.403]],
  '2_Custom Layout': [[11.775,6.243], [7.538,6.789]],
  'Picture with Caption': [[5.631,6.858]],
  'Title and Vertical Text': [[0.542,1.222]],
  '4_Custom Layout': [[12.422,6.243]],
  '5_Custom Layout': [[12.422,6.243]],
  '24_Custom Layout': [[3.792,6.858], [10.521,5.844]],
  '25_Custom Layout': [[0.765,4.605]],
  '26_Custom Layout': [[11.886,3.389]],
  '27_Custom Layout': [[0.748,4.741]],
  '28_Custom Layout': [[11.886,3.389]],
  '29_Custom Layout': [[4.049,6.858]],
  '31_Custom Layout': [[11.567,6.243]],
  '32_Custom Layout': [[4.874,1.091], [0.637,0.695]],
  '33_Custom Layout': [[4.049,6.858]],
  '18_Custom Layout': [[11.567,0.893]],
  '20_Custom Layout': [[4.049,3.083]],
};

// Circuit motifs per slide: [artIndex, x, y, color].
const DECOR = {
  1: [[0,9.935,1.081,LIME], [1,11.88,1.083,LIME], [2,10.636,0.381,LIME], [3,11.434,2.752,LIME], [4,11.88,2.22,LIME], [5,9.05,4.976,LIME], [6,10.624,5.314,LIME], [7,9.832,4.509,LIME], [8,7.157,6.12,LIME], [9,9.049,6.114,LIME], [10,5.785,7.08,LIME], [11,0.205,1.149,LIME], [75,1.851,4.729,LIME]],
  2: [[12,11.88,4.814,LIME], [5,9.05,4.976,LIME], [13,10.624,5.314,LIME], [8,7.157,6.12,LIME], [9,9.049,6.114,LIME], [10,5.785,7.08,LIME], [11,0.205,1.149,DARK], [76,11.931,3.529,DARK], [75,1.851,4.729,LIME]],
  3: [[14,11.877,3.808,LIME], [15,11.88,1.556,LIME], [16,11.877,5.5,LIME], [11,12.191,4.271,LIME], [17,11.877,3.022,LIME], [11,0.205,1.149,DARK], [76,5.019,6.867,DARK]],
  4: [[18,0.255,4.372,LIME], [19,0.546,2.152,LIME], [20,0.541,6.255,LIME], [11,12.597,1.149,DARK], [21,0.616,0.665,LIME], [76,7.327,6.867,DARK]],
  5: [[22,9.835,5.944,LIME], [8,7.157,6.12,LIME], [23,11.877,5.947,LIME], [9,9.049,6.348,LIME], [10,5.785,7.08,LIME], [24,4.061,6.442,LIME], [25,4.511,5.947,LIME], [26,1.484,6.442,LIME], [11,0.205,1.149,DARK], [76,5.668,5.09,DARK]],
  6: [[22,2.232,5.944,LIME], [27,2.288,7.088,LIME], [23,4.276,5.947,LIME], [28,1.446,6.348,LIME], [11,12.597,1.149,DARK], [76,6.314,6.867,DARK]],
  7: [[4,11.88,4.819,LIME], [5,9.05,4.976,LIME], [6,10.624,5.314,LIME], [8,7.157,6.12,LIME], [9,9.049,6.114,LIME], [10,5.785,7.08,LIME], [11,0.205,1.149,DARK], [76,11.931,3.529,DARK], [75,1.851,4.729,LIME]],
  8: [[11,0.205,1.149,DARK], [22,2.232,5.944,LIME], [27,2.288,7.088,LIME], [23,4.276,5.947,LIME], [29,0.634,6.26,LIME], [30,0.78,5.549,LIME], [10,3.591,5.558,LIME], [76,6.314,6.867,DARK]],
  9: [[18,0.255,4.372,LIME], [19,0.546,2.152,LIME], [20,0.541,6.255,LIME], [21,0.616,0.665,LIME], [11,12.597,1.149,DARK], [76,6.947,6.867,DARK]],
  10: [[11,0.205,1.149,DARK], [31,7.546,5.723,LIME], [27,9.188,7.088,LIME], [23,11.175,5.947,LIME], [32,7.656,6.348,LIME], [76,5.019,6.867,DARK]],
  11: [[11,0.205,1.149,DARK], [33,2.457,6.372,LIME], [34,4.932,6.369,LIME], [28,2.102,6.372,LIME], [35,5.781,6.761,LIME], [76,7.518,6.867,DARK]],
  12: [[11,12.597,1.149,DARK], [36,0.546,1.482,LIME], [21,0.616,0.484,LIME], [37,11.706,3.955,LIME], [20,11.992,6.321,LIME], [76,5.612,6.867,DARK]],
  13: [[11,0.205,1.149,DARK], [38,9.354,5.731,LIME], [27,9.835,7.088,LIME], [23,11.822,5.947,LIME], [28,8.992,6.348,LIME], [76,6.816,6.867,DARK]],
  14: [[39,10.261,4.975,LIME], [40,10.245,7.088,LIME], [23,11.822,5.947,LIME], [11,0.205,1.149,DARK], [76,8.675,6.867,DARK]],
  15: [[4,11.88,4.819,LIME], [5,9.05,4.976,LIME], [6,10.624,5.314,LIME], [8,7.157,6.12,LIME], [9,9.049,6.114,LIME], [10,5.785,7.08,LIME], [11,0.205,1.149,DARK], [76,11.931,3.529,DARK], [75,1.851,4.729,LIME]],
  16: [[41,10.527,4.125,LIME], [42,10.521,6.605,LIME], [43,0.205,0.374,LIME], [44,1.125,3.049,LIME], [45,0.39,2.855,LIME], [46,2.598,3.405,LIME], [47,3.935,6.798,LIME], [48,3.784,4.066,LIME], [49,5.891,7.157,LIME], [50,6.96,0.374,LIME]],
  17: [[51,0.447,0.499,LIME], [52,0.466,2.865,LIME], [53,0.768,3.16,LIME], [21,0.839,3.868,LIME], [11,12.693,1.149,DARK], [76,4.676,6.867,DARK]],
  18: [[54,9.805,0.499,LIME], [55,9.805,2.865,LIME], [56,10.195,3.868,LIME], [11,0.205,1.149,DARK], [76,8.2,6.867,DARK]],
  19: [[57,0.46,0.499,LIME], [58,0.46,2.865,LIME], [59,0.768,3.16,LIME], [21,0.839,3.868,LIME], [60,0.46,4.865,LIME], [11,12.693,1.149,DARK], [76,4.676,6.867,DARK]],
  20: [[61,9.805,0.499,LIME], [55,9.805,2.865,LIME], [56,10.195,3.868,LIME], [11,0.205,1.149,DARK], [76,8.2,6.867,DARK]],
  21: [[33,0.875,6.372,LIME], [34,3.351,6.369,LIME], [28,0.52,6.372,LIME], [35,4.2,6.761,LIME], [11,0.205,1.149,DARK], [62,5.785,6.495,LIME], [11,12.213,6.38,LIME], [63,9.312,6.372,LIME], [12,10.488,6.495,LIME], [64,5.785,6.8,LIME], [65,5.282,6.372,LIME], [66,5.777,6.51,LIME]],
  22: [[12,11.88,4.814,LIME], [5,9.05,4.976,LIME], [13,10.624,5.314,LIME], [8,7.157,6.12,LIME], [9,9.049,6.114,LIME], [10,5.785,7.08,LIME], [11,0.205,1.149,DARK], [76,11.931,3.529,DARK], [75,1.851,4.729,LIME]],
  23: [[11,0.205,1.149,DARK], [38,8.498,5.731,LIME], [27,8.98,7.088,LIME], [23,10.968,5.947,LIME], [28,8.139,6.348,LIME], [67,7.407,5.731,LIME], [76,5.716,6.867,DARK]],
  24: [[11,12.693,1.149,DARK], [22,2.232,0.792,LIME], [23,4.276,0.795,LIME], [9,1.446,1.196,LIME], [68,0.634,1.108,LIME], [30,0.78,0.397,LIME], [10,3.591,0.406,LIME], [76,12.041,6.867,DARK]],
  25: [[11,0.205,1.149,DARK], [33,0.875,6.372,LIME], [34,3.351,6.369,LIME], [28,0.52,6.372,LIME], [35,4.2,6.761,LIME], [62,5.785,6.495,LIME], [11,12.213,6.38,LIME], [63,9.312,6.372,LIME], [12,10.488,6.495,LIME], [64,5.785,6.8,LIME], [65,5.282,6.372,LIME], [66,5.777,6.51,LIME]],
  26: [[12,11.88,4.814,LIME], [5,9.05,4.976,LIME], [13,10.624,5.314,LIME], [8,7.157,6.12,LIME], [9,9.049,6.114,LIME], [10,5.785,7.08,LIME], [11,0.205,1.149,DARK], [76,11.931,3.529,DARK], [75,1.851,4.729,LIME]],
  27: [[11,0.205,1.149,DARK], [38,8.498,0.381,LIME], [69,8.937,0.805,LIME], [23,10.968,0.598,LIME], [28,8.139,0.997,LIME], [70,7.407,0.381,LIME], [71,9.117,2.1,LIME], [76,12.041,6.867,DARK]],
  28: [[72,11.78,2.577,LIME], [73,11.78,5.055,LIME], [11,12.399,3.825,LIME], [11,0.205,1.149,DARK], [74,11.777,4.606,LIME], [76,12.041,6.867,DARK]],
  29: [[11,0.205,1.149,DARK], [33,0.875,2.596,LIME], [34,3.351,2.594,LIME], [28,0.52,2.597,LIME], [35,4.2,2.986,LIME], [62,5.785,2.72,LIME], [11,12.213,2.604,LIME], [63,9.312,2.596,LIME], [12,10.488,2.72,LIME], [64,5.785,3.025,LIME], [65,5.282,2.596,LIME], [66,5.777,2.735,LIME], [76,12.041,6.867,DARK]],
  30: [[0,9.935,1.081,LIME], [1,11.88,1.083,LIME], [2,10.636,0.381,LIME], [3,11.434,2.752,LIME], [4,11.88,2.22,LIME], [5,9.05,4.976,LIME], [6,10.624,5.314,LIME], [7,9.832,4.509,LIME], [8,7.157,6.12,LIME], [9,9.049,6.114,LIME], [10,5.785,7.08,LIME], [11,0.205,1.149,LIME]],
};

// Which layout's panels/photos each slide inherits.
const LAYOUT_OF = {
  1: 'Title Slide', 2: '21_Custom Layout', 3: 'Title and Content',
  4: 'Section Header', 5: 'Two Content', 6: 'Comparison',
  7: '22_Custom Layout', 8: 'Blank', 9: 'Content with Caption',
  10: '2_Custom Layout', 11: 'Picture with Caption', 12: 'Title and Vertical Text',
  13: '4_Custom Layout', 14: '5_Custom Layout', 15: '23_Custom Layout',
  16: '24_Custom Layout', 17: '25_Custom Layout', 18: '26_Custom Layout',
  19: '27_Custom Layout', 20: '28_Custom Layout', 21: '29_Custom Layout',
  22: '30_Custom Layout', 23: '31_Custom Layout', 24: '32_Custom Layout',
  25: '33_Custom Layout', 26: '34_Custom Layout', 27: '18_Custom Layout',
  28: '19_Custom Layout', 29: '20_Custom Layout', 30: 'Title Slide',
};

/** Draw the layout chrome (panels + circuit art) for a slide. */
function chrome(s, num) {
  const lay = LAYOUT_OF[num];
  (PANELS[lay] || []).forEach(p => panel(s, p[0], p[1], p[2], p[3]));
  (ROSETTES[lay] || []).forEach(p => rosette(s, p[0], p[1]));
  (DECOR[num] || []).forEach(d => art(s, d[0], d[1], d[2], d[3]));
}

// ------------------------------------------------------------- slides

/** Slide 1 - Green */
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: DARK };
  chrome(s, 1);
  s.addText([
    { text: 'Green', options: { fontSize:72, fontFace:BLACK, color:LIME } },
    { text: ' Energy', options: { fontSize:72, fontFace:BLACK, color:CREAM } },
  ], { x:1.675, y:0.863, w:8.097, h:1.313, valign:'top', margin:MARGIN });
  tx(s, [1.641, 5.06, 5.32, 0.908], 'bodyLight', 'That way, the audience, who only has a little time, f listening  your presentation until the  end. For example, if you work in a company engaged in the property sector, you can');
  tx(s, [1.84, 2.446, 2.647, 0.707], 'labelLight', 'Presentation Template');
}

/** Slide 2 - Introduce */
function slide02(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 2);
  pageTag(s, 'Page - 02', DARK);
  tx(s, [1.641, 1.111, 4.364, 0.841], 'title', 'Introduce');
  tx(s, [1.641, 5.06, 5.32, 0.908], 'bodyLight', 'That way, the audience, who only has a little time, f listening  your presentation until the  end. For example, if you work in a company engaged in the property sector, you can');
  rule(s, 0, 3.528, 11.721, DARK);
}

/** Slide 3 - Welcome */
function slide03(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 3);
  tx(s, [1.59, 3.553, 4.274, 2.296], 'body', paras(['Us ally this type of opening presentation is used so that the audience immediately understands the problems that are  happening. ', '', 'That way, the audience, who only has a little time, will stay focused on listening to your presentation until engaged in the property sector, you can follow an example like this:']));
  rule(s, 0, 6.866, 4.75, DARK);
  tx(s, [1.641, 1.111, 4.364, 0.841], 'title', 'Welcome');
  pageTag(s, 'Page - 03', LIME);
}

/** Slide 4 - Our Story */
function slide04(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 4);
  tx(s, [7.447, 2.859, 4.925, 3.407], 'body', paras([':introduce our company, and we also hope that we can be partners in the company that you lead. As the beginning of production of herbal medicines which has been running for approximately 6 years.', '', 'Meanwhile, we have a superior product, namely but-but bird oil and have a turnover of more than 5,000 packs of production per day.', '', 'satisfaction. In addition, we also have a always be a supplier that each of our customers can rely on, and to make this company always number one in terms of innovation.']));
  tx(s, [7.447, 1.111, 4.364, 0.841], 'title', 'Our Story');
  rule(s, 8.292, 6.866, 5.042, DARK);
  pageTag(s, 'Page - 04', DARK);
}

/** Slide 5 - About Our Company */
function slide05(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 5);
  tx(s, [1.649, 3.124, 4.859, 1.185], 'body', ':introduce our company, and we also hope that we can be partners in the company that you lead. As the beginning of our production of herbal medicines which has been running for approximately 6 years.');
  tx(s, [1.636, 1.024, 4.364, 1.582], 'title', 'About Our Company');
  pageTag(s, 'Page - 05', DARK);
  rule(s, 0, 5.09, 5.467, DARK);
}

/** Slide 6 - About Our Business */
function slide06(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 6);
  tx(s, [6.499, 3.357, 5.659, 2.852], 'body', paras([':introduce our company, and we also hope that partners the company that you lead. As the beginning currently company is production herbal medicines which has been running approximately 6 years.', '', 'Meanwhile, we have a superior product, namely but-but bird oil and have a turnover of more than 5,000 packs of', '', 'satisfaction. In addition, we also have a always be a supplier that each of our customers can rely on, and to make this company number one in terms of innovation.']));
  tx(s, [6.499, 1.024, 4.364, 1.582], 'title', 'About Our Business');
  pageTag(s, 'Page - 06', DARK);
  rule(s, 7.257, 6.866, 6.076, DARK);
}

/** Slide 7 - Product */
function slide07(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 7);
  tx(s, [1.641, 1.111, 4.364, 0.841], 'title', 'Product');
  rule(s, 0, 3.528, 11.721, DARK);
  pageTag(s, 'Page - 07', DARK);
  tx(s, [1.641, 5.06, 5.32, 0.908], 'bodyLight', 'That way, the audience, who only has a little time, f listening  your presentation until the  end. For example, if you work in a company engaged in the property sector, you can');
}

/** Slide 8 - Our Product */
function slide08(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 8);
  tx(s, [1.637, 3.234, 4, 0.988], 'small', 'In manufacturing, products are purchased in the form of raw goods and sold as finished goods. Products in the form of raw goods such as metal or agricultural products are often referred to as commodities.');
  tx(s, [6.329, 4.473, 0.854, 0.64], 'stat', '01.');
  tx(s, [7.196, 4.647, 1.448, 0.404], 'label', 'Category');
  tx(s, [9.494, 4.473, 0.937, 0.64], 'stat', '02.');
  tx(s, [10.361, 4.647, 1.448, 0.404], 'label', 'Category');
  tx(s, [1.641, 1.111, 3.883, 1.582], 'title', 'Our Product');
  pageTag(s, 'Page - 08', DARK);
  rule(s, 7.257, 6.866, 6.076, DARK);
  tx(s, [6.287, 5.192, 2.735, 0.988], 'small', 'purchased in the form of raw goods and sold as finished goods. Products in form of raw goods such as products are often referred to as commodities.');
  tx(s, [9.493, 5.192, 2.735, 0.988], 'small', 'purchased in the form of raw goods and sold as finished goods. Products in form of raw goods such as products are often referred to as commodities.');
}

/** Slide 9 - Gallery */
function slide09(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 9);
  tx(s, [6.997, 1.111, 5.131, 0.841], 'title', 'Gallery');
  pageTag(s, 'Page - 09', DARK);
  rule(s, 7.927, 6.866, 5.406, DARK);
}

/** Slide 10 - Renewable Energy */
function slide10(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 10);
  tx(s, [1.641, 3.449, 4.463, 1.185], 'body', 'In manufacturing, products are purchased in the form of raw goods and sold as finished goods. Products in the form of raw goods such as metal or agricultural products are often referred to as commodities.');
  tx(s, [1.641, 5.429, 2.359, 0.908], 'body', 'In business, products are goods or services that can be traded. ');
  tx(s, [1.641, 5.025, 1.903, 0.404], 'label', 'Go Green');
  tx(s, [1.641, 1.111, 3.883, 1.582], 'title', 'Renewable Energy');
  rule(s, 0, 6.866, 4.75, DARK);
  pageTag(s, 'Page - 10', DARK);
}

/** Slide 11 - Reliable Energy */
function slide11(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 11);
  tx(s, [7.649, 5.059, 2.097, 0.908], 'body', 'In business, products are goods or services that can be traded. ');
  tx(s, [10.284, 5.059, 2.097, 0.908], 'body', 'In business, products are goods or services that can be traded');
  tx(s, [7.649, 2.289, 4.732, 1.185], 'body', 'In manufacturing, products are purchased in the form of raw goods and sold as finished goods. Products in the form of raw goods such as metal or agricultural products are often referred to as commodities.');
  tx(s, [1.641, 1.111, 5.519, 0.841], 'title', 'Reliable Energy');
  tx(s, [7.693, 3.944, 0.937, 0.64], 'stat', '01.');
  tx(s, [7.693, 4.584, 1.448, 0.404], 'label', 'Category');
  tx(s, [10.381, 3.944, 0.937, 0.64], 'stat', '02.');
  tx(s, [10.381, 4.584, 1.448, 0.404], 'label', 'Category');
  rule(s, 8.542, 6.866, 4.792, DARK);
  pageTag(s, 'Page - 11', DARK);
}

/** Slide 12 - Comparison Energy */
function slide12(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 12);
  tx(s, [2.204, 5.57, 1.796, 0.908], 'body', 'products are goods or services that can be traded. ');
  tx(s, [4.449, 5.57, 1.796, 0.908], 'body', 'products are goods or services that can be traded');
  tx(s, [7.111, 1.111, 4.665, 1.582], 'title', 'Comparison Energy');
  tx(s, [2.204, 4.396, 1.362, 0.64], 'stat', '65%');
  tx(s, [2.204, 5.036, 1.796, 0.404], 'label', 'Type A');
  tx(s, [4.517, 4.396, 1.236, 0.64], 'stat', '80%');
  tx(s, [4.517, 5.036, 1.728, 0.404], 'label', 'Type B');
  rule(s, 0, 6.866, 5.347, DARK);
  pageTag(s, 'Page - 12', DARK);
}

/** Slide 13 - Interest in Use */
function slide13(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 13);
  tx(s, [1.643, 2.279, 1.903, 0.64], 'stat', '75%');
  tx(s, [5.015, 2.918, 1.903, 0.404], 'label', 'Electricity');
  tx(s, [5.015, 2.305, 1.903, 0.64], 'stat', '62%');
  tx(s, [1.643, 2.918, 1.903, 0.404], 'label', 'Crude Oil');
  tx(s, [8.421, 2.295, 3.96, 1.185], 'body', 'In manufacturing, products are purchased in the form of raw goods and sold as finished Products in the form of raw goods such as products are often referred to as commodities.');
  tx(s, [8.421, 3.951, 1.903, 0.404], 'label', 'Conclusion');
  tx(s, [8.421, 4.366, 3.96, 0.63], 'body', 'In manufacturing, products are purchased in the form of raw goods and sold as finished goods.');
  pageTag(s, 'Page - 13', DARK);
  tx(s, [1.641, 1.111, 5.519, 0.841], 'title', 'Interest in Use');
  rule(s, 0, 6.866, 6.571, DARK);
}

/** Slide 14 - Alternative Energy */
function slide14(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 14);
  tx(s, [1.641, 3.309, 4.728, 0.908], 'body', 'In manufacturing, products are purchased in the form of raw raw goods such as metal or agricultural products are often referred to as commodities.');
  tx(s, [1.641, 4.862, 1.903, 0.404], 'label', 'Options');
  tx(s, [1.641, 5.342, 3.206, 0.908], 'body', 'In manufacturing, products purchased in the form of raw goods and sold as finished goods.');
  tx(s, [1.641, 1.111, 5.519, 1.582], 'title', 'Alternative Energy');
  pageTag(s, 'Page - 14', DARK);
  rule(s, 0, 6.866, 8.513, DARK);
}

/** Slide 15 - Analysis */
function slide15(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 15);
  tx(s, [1.641, 1.111, 4.364, 0.841], 'title', 'Analysis');
  rule(s, 0, 3.528, 11.721, DARK);
  pageTag(s, 'Page - 15', DARK);
  tx(s, [1.641, 5.06, 5.32, 0.908], 'bodyLight', 'That way, the audience, who only has a little time, f listening  your presentation until the  end. For example, if you work in a company engaged in the property sector, you can');
}

/** Slide 16 */
function slide16(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 16);
  tx(s, [1.381, 0.378, 1.506, 2.625], 'letter', 'S', { align:'center', wrap:false });
  tx(s, [11.541, 4.128, 1.506, 2.625], 'letter', 'T', { align:'center', wrap:false });
  tx(s, [0.752, 4.231, 1.977, 0.404], 'label', 'Strengths', { align:'center', charSpacing:1 });
  tx(s, [0.997, 4.813, 1.487, 1.185], 'body', 'component of SWOT is strength or power in the business.', { align:'center' });
  tx(s, [3.73, 1.532, 2.511, 0.404], 'label', 'Weaknesses', { align:'center', charSpacing:1 });
  tx(s, [3.982, 2.114, 1.998, 1.185], 'body', 'In SWOT analysis, W is a weakness which means the weakness of the company or', { align:'center' });
  tx(s, [7.203, 4.231, 2.276, 0.404], 'label', 'Opportunities', { align:'center', charSpacing:1 });
  tx(s, [7.375, 4.813, 1.932, 1.185], 'body', 'The next SWOT opportunities, which means business opportunities.', { align:'center' });
  tx(s, [10.49, 1.646, 2.047, 0.404], 'label', 'Threats', { align:'center', charSpacing:1 });
  tx(s, [10.595, 2.228, 1.838, 1.185], 'body', 'Meanwhile, the SWOT analysis related to business threats is threats.', { align:'center' });
  pageTag(s, 'Page - 16', DARK);
  tx(s, [4.076, 4.128, 2.465, 2.625], 'letter', 'W', { align:'center', wrap:false });
  tx(s, [8.003, 0.378, 1.857, 2.625], 'letter', 'O', { align:'center', wrap:false });
}

/** Slide 17 - Strengths */
function slide17(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 17);
  tx(s, [4.621, 4.577, 7.046, 1.463], 'body', bullets(['What has your business been doing well?', 'What do consumers say about the benefits of the product/service?', 'What are the advantages of your business product/service?', 'What is the advantage of your business over competitors?', 'What assets stand out in your business over competitors?']));
  tx(s, [4.621, 2.424, 7.252, 1.741], 'body', paras(['Strength in SWOT analysis is the strength or advantage of internal factors that your business has compared to competitors. For example, business management, employee performance, products created, and various uniqueness built by your own business.', '', 'Because it focuses on the advantages, uniqueness, and strengths that your business internal factors have, here are the questions that can be your reference in conducting']));
  tx(s, [4.621, 1.111, 4.364, 0.841], 'title', 'Strengths');
  pageTag(s, 'Page - 17', DARK);
  tx(s, [1.966, 0.378, 1.506, 2.625], 'letter', 'S', { align:'center', wrap:false });
  rule(s, 5.683, 6.866, 7.651, DARK);
}

/** Slide 18 - Weaknesses */
function slide18(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 18);
  tx(s, [1.641, 2.376, 6.606, 2.296], 'body', paras(['No business is always perfect. You must have experienced various problems, and in the they were covered by the advantages and uniqueness that make people love your business.', '', 'This is the importance of analyzing the weakness elements in a SWOT analysis for your business development by relying on internal factors, so that these weaknesses can be overcome. Finally, your business creates various strategies with innovations and crazy ideas.']));
  tx(s, [1.641, 4.993, 7.046, 1.185], 'body', bullets(['What can your business improve?', 'What makes customers dissatisfied with your business?', 'Where does your business lag behind competitors?', 'What knowledge or resources does your business lag behind the competition?']));
  tx(s, [10.633, 4.384, 2.251, 2.625], 'letter', 'W', { align:'center' });
  tx(s, [1.641, 1.111, 4.364, 0.841], 'title', 'Weaknesses');
  rule(s, 0, 6.866, 7.967, DARK);
  pageTag(s, 'Page - 18', LIME);
}

/** Slide 19 - Opportunities */
function slide19(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 19);
  tx(s, [4.621, 2.771, 6.606, 1.741], 'body', paras(['Remember! A successful business is a business that is able to survive because it wants to take various opportunities from what you often encounter.', '', 'This is an example of an opportunity element in a SWOT analysis. Then, how do you describe various opportunities from external factors into sticky notes that will be pasted on the board? Roughly the question is as follows:']));
  tx(s, [4.621, 4.807, 7.046, 1.185], 'body', bullets(['What trends can your business follow?', 'What strengths can add value to our business in the eyes of consumers?', 'How is the competition around your business location?', 'What opportunities can your business take advantage of?']));
  tx(s, [1.791, 0.378, 1.857, 2.625], 'letter', 'O', { align:'center', wrap:false });
  rule(s, 5.683, 6.866, 7.651, DARK);
  tx(s, [4.621, 1.111, 5.554, 0.841], 'title', 'Opportunities');
  pageTag(s, 'Page - 19', DARK);
}

/** Slide 20 - Threats */
function slide20(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 20);
  tx(s, [1.641, 2.553, 6.553, 2.019], 'body', paras(['Do you feel that this pandemic period is really hard for a business that has an offline store? Such as shops in malls, restaurants, tourist attractions, hotels.', '', 'Since the pandemic, offline-based businesses such as the one above have had to limit operating hours, limit visiting consumers, and even temporarily close due to government regulations to prevent the spread of the Covid-19 virus. The limitations of business activities make their incomes decrease, even many who lose. Pity, huh?']));
  tx(s, [1.641, 5.02, 7.046, 1.185], 'body', bullets(['What are competitors doing?', 'What market trends is your business missing?', 'What economic or political issues affect your business?', 'What are the weaknesses of your business that make it vulnerable?']));
  tx(s, [10.633, 4.384, 2.251, 2.625], 'letter', 'T', { align:'center' });
  tx(s, [1.641, 1.111, 4.364, 0.841], 'title', 'Threats');
  rule(s, 0, 6.866, 7.967, DARK);
  pageTag(s, 'Page - 20', LIME);
}

/** Slide 21 - 55% */
function slide21(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 21);
  s.addShape('rect', { x:5.665, y:3.028, w:2.002, h:2.002, fill:{ type:'none' }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:7.113, y:2.435, w:1.67, h:1.421, fill:{ color:DARK } });
  s.addShape('rect', { x:7.113, y:4.24, w:1.67, h:1.421, fill:{ color:DARK } });
  s.addShape('rect', { x:4.55, y:2.435, w:1.67, h:1.421, fill:{ color:DARK } });
  s.addShape('rect', { x:4.55, y:4.24, w:1.67, h:1.421, fill:{ color:DARK } });
  tx(s, [1.932, 2.44, 1.977, 0.404], 'label', 'Strengths', { align:'right', charSpacing:1 });
  tx(s, [1.457, 2.869, 2.452, 0.908], 'body', 'The first component of SWOT is strength or power in the business.', { align:'right' });
  tx(s, [4.745, 2.71, 1.306, 0.841], 'pct', '55%', { align:'center' });
  tx(s, [1.399, 4.272, 2.511, 0.404], 'label', 'Weaknesses', { align:'right', charSpacing:1 });
  tx(s, [1.307, 4.701, 2.603, 0.761], 'small', 'In SWOT analysis, W is a weakness which means the weakness of the company or business.', { align:'right' });
  tx(s, [9.404, 2.44, 2.389, 0.404], 'label', 'Opportunities', { charSpacing:1 });
  tx(s, [9.404, 2.869, 2.598, 0.908], 'body', 'The next SWOT component is opportunities, which means business opportunities.');
  tx(s, [9.404, 4.272, 2.047, 0.404], 'label', 'Threats', { charSpacing:1 });
  tx(s, [9.404, 4.701, 2.334, 0.761], 'small', 'Meanwhile, the SWOT analysis related to business threats is threats.');
  tx(s, [4.745, 4.544, 1.306, 0.841], 'pct', '5%', { align:'center' });
  tx(s, [7.254, 2.71, 1.306, 0.841], 'pct', '25%', { align:'center' });
  tx(s, [7.254, 4.544, 1.306, 0.841], 'pct', '15%', { align:'center' });
  tx(s, [3.884, 1.111, 5.554, 0.841], 'title', 'Percentages', { align:'center' });
  pageTag(s, 'Page - 21', DARK);
}

/** Slide 22 - Infographic */
function slide22(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 22);
  tx(s, [1.641, 1.111, 4.364, 0.841], 'title', 'Infographic');
  rule(s, 0, 3.528, 11.721, DARK);
  pageTag(s, 'Page - 22', DARK);
  tx(s, [1.641, 5.06, 5.32, 0.908], 'bodyLight', 'That way, the audience, who only has a little time, f listening  your presentation until the  end. For example, if you work in a company engaged in the property sector, you can');
}

/** Slide 23 - Product Chart */
function slide23(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 23);
  // Stock (open-high-low-close) chart, five sessions.
  const OHLC = [
    { name: 'Open',  labels: ['1', '2', '3', '4', '5'], values: [44, 25, 38, 50, 34] },
    { name: 'High',  labels: ['1', '2', '3', '4', '5'], values: [55, 57, 57, 58, 36] },
    { name: 'Low',   labels: ['1', '2', '3', '4', '5'], values: [11, 12, 13, 11,  5] },
    { name: 'Close', labels: ['1', '2', '3', '4', '5'], values: [25, 38, 50, 34, 18] },
  ];
  s.addChart('line', OHLC, {
    x: 7.014, y: 1.107, w: 5.297, h: 3.819,
    chartColors: [LIME, DARK, '595959', LIME],
    lineSize: 1, showLegend: false, showValue: false,
    catAxisHidden: true, valAxisLineColor: DARK,
    valGridLine: { color: LIME, style: 'solid', size: 1 },
    valAxisLabelColor: '595959', valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12,
    plotArea: { border: { color: DARK, pt: 1 }, fill: { color: CREAM } },
  });
  tx(s, [1.648, 3.108, 4.462, 1.185], 'body', 'In manufacturing, products are purchased in the form of raw goods and sold as finished goods. Products in the form of raw goods such as metal or agricultural products are often referred to as commodities.');
  tx(s, [1.639, 5.307, 4.471, 0.908], 'body', 'In business, products are goods or services that can be traded. In marketing, a product is anything that can be offered to a market and can satisfy a want or need.');
  tx(s, [1.639, 4.807, 1.794, 0.404], 'label', 'All Category');
  tx(s, [1.641, 1.111, 4.462, 1.582], 'title', 'Product Chart');
  rule(s, 0, 6.866, 5.573, DARK);
  pageTag(s, 'Page - 23', DARK);
}

/** Slide 24 - Energy Chart */
function slide24(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 24);
  // Histogram of 76 samples binned into six equal intervals over 1..24.
  s.addChart('bar', [{
    name: 'Series1',
    labels: ['1-5', '5-9', '9-13', '13-17', '17-21', '21-24'],
    values: [5, 11, 23, 24, 9, 4],
  }], {
    x: 0.904, y: 2.709, w: 4.706, h: 3.565,
    barDir: 'col', barGapWidthPct: 0,
    chartColors: [LIME], chartColorsOpacity: 100,
    dataBorder: { pt: 0.75, color: '595959' },
    showLegend: false, showValue: false,
    catAxisHidden: true, catAxisLineShow: false,
    valAxisLineColor: DARK, valAxisMaxVal: 25, valAxisMajorUnit: 5,
    valGridLine: { color: DARK, style: 'solid', size: 1 },
    valAxisLabelColor: DARK, valAxisLabelFontFace: BODY, valAxisLabelFontSize: 12,
    plotArea: { fill: { color: CREAM } },
  });
  tx(s, [6.799, 1.111, 4.462, 0.841], 'title', 'Energy Chart');
  tx(s, [6.799, 2.71, 4.889, 1.185], 'body', 'In manufacturing, products are purchased in the form of raw goods and sold as finished goods. Products in the form of raw goods such as metal or agricultural products often referred to as commodities.');
  tx(s, [6.799, 5.263, 5.042, 0.908], 'body', 'In business, products are goods or services that can be traded. In marketing, a product is anything that can be offered to be a market and can satisfy a want or need.');
  tx(s, [6.799, 4.747, 1.794, 0.404], 'label', 'All Category');
  rule(s, 0, 6.866, 11.841, DARK);
  pageTag(s, 'Page - 24', DARK);
}

/** Slide 25 - Comparison Chart */
function slide25(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 25);
  s.addShape('rect', { x:1.632, y:3.851, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:2.551, y:3.851, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:3.47, y:3.851, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:4.388, y:3.851, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:5.307, y:3.851, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:6.225, y:3.851, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:7.144, y:3.851, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:8.062, y:3.851, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:8.981, y:3.851, w:0.708, h:0.265, fill:{ color:DARK }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:9.899, y:3.851, w:0.708, h:0.265, fill:{ color:DARK }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:1.632, y:5.053, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:2.551, y:5.053, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:3.47, y:5.053, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:4.388, y:5.053, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:5.307, y:5.053, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:6.225, y:5.053, w:0.708, h:0.265, fill:{ color:LIME }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:7.144, y:5.053, w:0.708, h:0.265, fill:{ color:DARK }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:8.062, y:5.053, w:0.708, h:0.265, fill:{ color:DARK }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:8.981, y:5.053, w:0.708, h:0.265, fill:{ color:DARK }, line:{ color:DARK, width:1 } });
  s.addShape('rect', { x:9.899, y:5.053, w:0.708, h:0.265, fill:{ color:DARK }, line:{ color:DARK, width:1 } });
  tx(s, [1.641, 1.111, 4.734, 1.582], 'title', 'Comparison Chart');
  tx(s, [1.54, 4.485, 3.314, 0.404], 'label', 'Category 02');
  tx(s, [1.54, 3.249, 3.314, 0.404], 'label', 'Category 01');
  tx(s, [10.844, 5, 0.857, 0.404], 'label', '60%', { align:'right' });
  tx(s, [10.844, 3.763, 0.857, 0.404], 'label', '80%', { align:'right' });
  tx(s, [7.681, 1.309, 4.012, 1.185], 'body', 'In manufacturing, products are purchased in the form of raw goods and sold as finished Products in the form of raw goods such as products are often referred to as commodities.');
  pageTag(s, 'Page - 25', DARK);
}

/** Slide 26 - Team */
function slide26(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 26);
  tx(s, [1.641, 1.111, 4.364, 0.841], 'title', 'Team');
  rule(s, 0, 3.528, 11.721, DARK);
  pageTag(s, 'Page - 26', DARK);
  tx(s, [1.641, 5.06, 5.32, 0.908], 'bodyLight', 'That way, the audience, who only has a little time, f listening  your presentation until the  end. For example, if you work in a company engaged in the property sector, you can');
}

/** Slide 27 - Team Performance */
function slide27(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 27);
  tx(s, [3.99, 3.363, 1.979, 0.707], 'label', 'Employee of the Month');
  tx(s, [3.99, 4.435, 1.974, 0.908], 'body', 'In business, products are goods or services that can be traded.');
  tx(s, [1.539, 5.631, 1.979, 0.404], 'label', 'Britney Ellie', { align:'center' });
  s.addShape('star5', { x:1.523, y:6.057, w:0.332, h:0.332, fill:{ type:'none' }, line:{ color:DARK, width:1 } });
  s.addShape('star5', { x:1.943, y:6.057, w:0.332, h:0.332, fill:{ type:'none' }, line:{ color:DARK, width:1 } });
  s.addShape('star5', { x:2.363, y:6.057, w:0.332, h:0.332, fill:{ type:'none' }, line:{ color:DARK, width:1 } });
  s.addShape('star5', { x:2.782, y:6.057, w:0.332, h:0.332, fill:{ type:'none' }, line:{ color:DARK, width:1 } });
  s.addShape('star5', { x:3.202, y:6.057, w:0.332, h:0.332, fill:{ color:DARK }, line:{ color:DARK, width:1 } });
  tx(s, [9.607, 3.363, 1.979, 0.707], 'label', 'Best Leadership');
  tx(s, [9.607, 4.435, 1.974, 0.908], 'body', 'In business, products are goods or services that can be traded.');
  tx(s, [7.177, 5.631, 1.979, 0.404], 'label', 'Sean Davis', { align:'center' });
  s.addShape('star5', { x:7.161, y:6.057, w:0.332, h:0.332, fill:{ type:'none' }, line:{ color:DARK, width:1 } });
  s.addShape('star5', { x:7.581, y:6.057, w:0.332, h:0.332, fill:{ type:'none' }, line:{ color:DARK, width:1 } });
  s.addShape('star5', { x:8.001, y:6.057, w:0.332, h:0.332, fill:{ type:'none' }, line:{ color:DARK, width:1 } });
  s.addShape('star5', { x:8.42, y:6.057, w:0.332, h:0.332, fill:{ color:DARK }, line:{ color:DARK, width:1 } });
  s.addShape('star5', { x:8.84, y:6.057, w:0.332, h:0.332, fill:{ color:DARK }, line:{ color:DARK, width:1 } });
  tx(s, [1.641, 1.111, 5.229, 1.582], 'title', 'Team Performance');
  pageTag(s, 'Page - 27', LIME);
  rule(s, 0, 6.866, 11.841, DARK);
}

/** Slide 28 - CEO Founder */
function slide28(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 28);
  s.addShape('rect', { x:11.387, y:2.163, w:1.946, h:4.003, fill:{ color:DARK } });
  tx(s, [1.523, 2.336, 5.667, 0.908], 'body', 'The Chief Executive Officer, which stands for CEO, is the highest position in a company and is responsible for the success of the business being run. Quoted from Investopedia, CEOs are generally chosen');
  tx(s, [1.523, 3.626, 3.046, 0.404], 'label', 'Michael J. Robinson');
  tx(s, [1.523, 4.148, 3.768, 2.019], 'body', bullets(['Leadership', 'Firm and able to make quick decisions', 'Good communication', 'Creativity and innovation', 'Ethics', 'Collaboration', 'Transparency']));
  tx(s, [1.641, 1.111, 5.229, 0.841], 'title', 'CEO Founder');
  rule(s, 0, 6.866, 11.841, DARK);
  pageTag(s, 'Page - 28', DARK);
}

/** Slide 29 - Keep in Touch */
function slide29(pptx) {
  const s = pptx.addSlide();
  s.background = { color: CREAM };
  chrome(s, 29);
  tx(s, [7.291, 4.021, 1.974, 0.404], 'label', 'Call Us');
  tx(s, [7.291, 4.498, 1.974, 0.63], 'body', paras(['+123 456 789000', '+123 634 234564']));
  tx(s, [7.291, 5.406, 1.974, 0.404], 'label', 'Email');
  tx(s, [7.291, 5.883, 1.974, 0.63], 'body', paras(['example@mail.com', 'Style_list@mail.com ']));
  tx(s, [9.697, 4.021, 2.16, 0.404], 'label', 'Web & social');
  tx(s, [9.697, 4.498, 2.16, 0.63], 'body', paras(['www.exampleweb.com ', '@example_2025']));
  tx(s, [9.697, 5.406, 1.974, 0.404], 'label', 'Address');
  s.addText([
    { text: '21', options: { fontSize:11, fontFace:BODY, color:DARK } },
    { text: 'st', options: { fontSize:11, fontFace:BODY, color:DARK, superscript:true } },
    { text: ' Main street Philadelphia US 432667', options: { fontSize:11, fontFace:BODY, color:DARK } },
  ], { x:9.697, y:5.883, w:1.974, h:0.63, lineSpacingMultiple:1.5, valign:'top', margin:MARGIN });
  tx(s, [1.641, 1.111, 5.229, 0.841], 'title', 'Keep in Touch');
  rule(s, 0, 6.866, 11.841, DARK);
  tx(s, [7.291, 1.116, 4.565, 0.908], 'body', 'The Chief Executive Officer, which stands for CEO, is the highest position in a company and is responsible for the success of the business being run. Quoted from');
  pageTag(s, 'Page - 29', DARK);
}

/** Slide 30 - Thank */
function slide30(pptx) {
  const s = pptx.addSlide();
  s.background = { color: DARK };
  chrome(s, 30);
  pageTag(s, 'Page - 30', LIME);
  s.addText([
    { text: 'Thank ', options: { fontSize:72, fontFace:BLACK, color:LIME } },
    { text: 'You', options: { fontSize:72, fontFace:BLACK, color:CREAM } },
  ], { x:1.523, y:0.863, w:5.325, h:2.524, valign:'top', margin:MARGIN });
  tx(s, [1.523, 4.603, 2.038, 0.404], 'labelLight', 'Composed by');
  tx(s, [1.523, 5.225, 2.429, 1.185], 'bodyLight', bullets(['Lorem Ipsum', 'Lorem Ipsum', 'Lorem Ipsum', 'Lorem Ipsum']));
}

// --------------------------------------------------------------- build
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.author = 'Green Energy template';
  pptx.title = 'Green Energy';

  [
    slide01, slide02, slide03, slide04, slide05, slide06,
    slide07, slide08, slide09, slide10, slide11, slide12,
    slide13, slide14, slide15, slide16, slide17, slide18,
    slide19, slide20, slide21, slide22, slide23, slide24,
    slide25, slide26, slide27, slide28, slide29, slide30,
  ].forEach(fn => fn(pptx));

  return pptx.writeFile({
    fileName: path.join(__dirname, '01d4a1dd-1243-4499-afce-29d342627924_grok_final.pptx'),
  });
}

build().then(f => console.log('wrote', f)).catch(e => { console.error(e); process.exit(1); });
