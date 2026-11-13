/**
 * ORNAMENT — Presentation Template  (50 slides, 13.333in x 7.5in)
 *
 * Standalone pptxgenjs re-creation of the reference deck.
 * Run:  node 0e07a616-0c2f-4605-9b29-085cbac6e66a_grok_final.js
 *
 * Design notes
 *  - Typography: "Staatliches" for display headings, "Open Sans" for body copy.
 *  - Palette: warm-grey ink on an off-white F2F2F2 canvas (see PALETTE below).
 *  - The reference deck's photographs are replaced by flat grey placeholder
 *    rectangles labelled "[image]" (see `img()`), per the no-raster-assets rule.
 *  - Vector icons / infographic silhouettes are re-drawn as custom geometry from
 *    the normalised outlines listed in the PATHS table.
 */

'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

/* ------------------------------------------------------------------ palette */

const INK = '404040';        // primary text / dark blocks
const MUTED = '595959';      // hairlines + body copy
const HEADLINE = '333333';   // a few slightly darker headings
const CANVAS = 'F2F2F2';     // slide background
const PAPER = 'FFFFFF';      // pure white (reversed text)
const PLACEHOLDER_FILL = 'DCDCDC';   // stands in for a photograph
const PLACEHOLDER_TEXT = '8A8A8A';

const DISPLAY = 'Staatliches';
const TEXT = 'Open Sans';

const HAIRLINE = { color: MUTED, width: 2.25 };
const NO_LINE = { type: 'none' };
const T = {
  orn: 'Ornament',
  p1: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo.',
  p10: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto',
  p11: 'Sed a lacinia nulla, vitae imperdiet dui. Donec iaculis nulla eu metus viverra, eget lacinia risus lobortis. Praesent in nulla eu dolor maximus semper. Nunc accumsan, nibh quis eleifend sodales, diam sapien rhoncus ligula, eu vulputate purus erat vitae lorem.',
  p2: 'Lorem ipsum dolor sit amet, consectetur',
  p3: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore',
  p4: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem',
  p5: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae',
  p6: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nam quis varius tortor. Pellentesque auctor aliquam lacus quis',
  p7: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nam quis varius tortor. Pellentesque auctor',
  p8: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, ',
  p9: 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam',
  sub: 'Your Amazing  Subheading Here',
  tag: 'Your Amazing  Tagline Here',
};

const CORNERS = {
  C1: { line: [.662, 0, 0, 1.477], text: [.123, 2.06, 1.077, .404], rotate: 270, align: 'right' },
  C2: { line: [11.857, .747, 1.477, 0], text: [10.292, .532, 1.46, .404], align: 'right' },
  C3: { line: [12.519, 0, 0, 1.477], text: [11.801, 2.109, 1.46, .404], rotate: 270, align: 'right' },
  C4: { line: [.682, 0, 0, 1.477], text: [-.035, 2.109, 1.46, .404], rotate: 270, align: 'right' },
  C5: { line: [.666, 0, 0, 1.477], text: [-.051, 2.109, 1.46, .404], rotate: 270, align: 'right' },
  C6: { line: [12.667, 6.023, 0, 1.477], text: [11.949, 4.987, 1.46, .404], rotate: 270, flipH: true, align: 'right' },
  C7: { line: [-.001, .77, 1.477, 0], text: [1.687, .555, 1.46, .404], flipH: true },
  C8: { line: [.65, 6.023, 0, 1.477], text: [-.067, 4.987, 1.46, .404], rotate: 270, flipH: true },
  C9: { line: [12.666, 0, 0, 1.477], text: [11.949, 2.109, 1.46, .404], rotate: 270, align: 'right' },
  C10: { line: [11.857, .749, 1.477, 0], text: [10.292, .559, 1.46, .404], align: 'right' },
  C11: { line: [11.857, 6.764, 1.477, 0], text: [10.292, 6.575, 1.46, .404], align: 'right' },
  C12: { line: [-.001, .757, 1.477, 0], text: [1.581, .567, 1.46, .404], rotate: 180, flipV: true },
  C13: { line: [-.001, 6.78, 1.477, 0], text: [1.581, 6.591, 1.46, .404], flipH: true },
  C14: { line: [.665, 0, 0, 1.477], text: [-.053, 2.109, 1.46, .404], rotate: 270, align: 'right' },
  C15: { line: [.65, 6.023, 0, 1.477], text: [-.421, 4.632, 2.169, .404], rotate: 270, flipH: true },
  C16: { line: [-.001, 6.76, 1.477, 0], text: [1.581, 6.545, 1.46, .404], flipH: true },
  C17: { line: [12.519, 0, 0, 1.477], text: [11.892, 2.019, 1.279, .404], rotate: 270, align: 'right' },
  C18: { line: [11.857, 6.76, 1.477, 0], text: [10.292, 6.545, 1.46, .404], align: 'right' },
  C19: { line: [11.857, .777, 1.477, 0], text: [10.292, .562, 1.46, .404], align: 'right' },
};

const P = {
  flag: 'M.99 .38C.47 1 .64 .37 .21 .67L.32 1L.2 1L0 .33L.1 .3C.59 0 .34 .56 .96 .36C.98 .35 1 .36 .99 .38Z',
  search: 'M.92 1C.9 1 .88 1 .87 .98L.66 .78C.6 .83 .51 .84 .42 .84C.19 .84 0 .66 0 .43C0 .18 .19 0 .42 0C.65 0 .85 .18 .85 .43C.85 .5 .82 .59 .77 .66L.98 .87C.99 .88 1 .9 1 .92C1 .97 .96 1 .92 1ZM.42 .16C.28 .16 .15 .27 .15 .43C.15 .57 .28 .69 .42 .69C.58 .69 .69 .57 .69 .43C.69 .27 .58 .16 .42 .16ZM.62 .45C.62 .46 .61 .47 .6 .47L.46 .47L.46 .59C.46 .61 .45 .61 .44 .61L.4 .61C.39 .61 .38 .61 .38 .59L.38 .47L.25 .47C.24 .47 .23 .46 .23 .45L.23 .41C.23 .39 .24 .39 .25 .39L.38 .39L.38 .25C.38 .23 .39 .23 .4 .23L.44 .23C.45 .23 .46 .23 .46 .25L.46 .39L.6 .39C.61 .39 .62 .39 .62 .41L.62 .45Z',
  hammer: 'M.08 .96C.08 .96 .07 .96 .06 .95C.05 .95 .04 .94 .04 .94C.02 .94 .02 .93 .01 .93C0 .92 0 .91 0 .91C0 .9 0 .9 0 .89C.01 .88 .01 .87 .02 .86C.04 .85 .04 .84 .05 .84C.18 .66 .29 .51 .38 .4C.38 .4 .38 .4 .39 .4L.39 .39L.4 .39C.42 .37 .44 .34 .46 .3C.49 .26 .51 .23 .52 .21C.52 .2 .52 .2 .54 .19C.54 .18 .55 .18 .55 .17C.56 .16 .56 .15 .56 .15C.55 .12 .52 .11 .49 .1C.46 .09 .44 .09 .4 .09C.39 .09 .37 .09 .36 .09C.31 .1 .26 .11 .23 .12C.23 .12 .21 .11 .23 .11L.23 .1C.23 .1 .23 .1 .24 .1L.24 .09C.23 .09 .21 .09 .23 .07L.24 .06L.25 .06C.26 .05 .26 .05 .27 .05C.27 .05 .29 .05 .29 .04C.3 .04 .3 .04 .31 .04L.32 .03L.35 .03C.4 .01 .45 0 .51 0C.61 0 .69 .03 .79 .07C.8 .09 .81 .11 .82 .12C.82 .13 .85 .14 .87 .14C.87 .14 .88 .14 .9 .14C.92 .14 .93 .14 .94 .14C.95 .15 .96 .15 .98 .16C.99 .17 1 .17 1 .17C1 .17 1 .17 1 .18C1 .18 1 .19 1 .2L.96 .25C.95 .27 .94 .27 .93 .28C.93 .28 .92 .28 .9 .27C.88 .26 .87 .25 .87 .25C.86 .24 .86 .23 .85 .22C.83 .21 .83 .2 .82 .2C.81 .18 .79 .17 .75 .17C.74 .17 .73 .18 .7 .18L.67 .19C.67 .2 .67 .2 .65 .2C.65 .21 .64 .21 .64 .22L.63 .23C.63 .24 .62 .25 .62 .25C.61 .27 .58 .3 .56 .35C.54 .39 .51 .42 .5 .44L.51 .44C.51 .44 .51 .45 .5 .45L.5 .46C.44 .58 .35 .74 .23 .93C.23 .93 .23 .94 .21 .95C.21 .96 .21 .97 .2 .98C.2 .99 .19 .99 .19 1L.18 1C.17 1 .17 1 .15 1C.14 .99 .13 .99 .12 .98C.1 .97 .08 .96 .08 .96Z',
  palette: 'M.92 .61C.92 .61 .76 .56 .76 .47C.76 .46 .76 .46 .76 .45C.7 .49 .63 .53 .61 .54L.58 .56L.57 .52C.56 .49 .55 .33 .58 .27C.6 .24 .62 .22 .66 .21L.67 .21L.67 .2C.67 .17 .67 .12 .68 .09L.7 .06L.7 .07L.72 .04C.72 .04 .72 .04 .73 .03C.56 0 .37 .05 .23 .18C.02 .36 0 .66 .19 .83C.38 1 .7 .98 .91 .8C.96 .75 .97 .72 .97 .72C1 .68 .98 .63 .92 .61ZM.32 .65C.29 .68 .24 .68 .21 .65C.18 .62 .18 .58 .21 .55C.24 .52 .29 .52 .32 .55C.35 .58 .35 .62 .32 .65ZM.25 .42C.22 .39 .22 .35 .25 .32C.28 .29 .33 .29 .36 .32C.39 .35 .39 .39 .36 .42C.33 .45 .28 .45 .25 .42ZM.51 .81C.48 .84 .43 .84 .4 .81C.37 .78 .37 .74 .4 .71C.43 .69 .48 .69 .51 .71C.54 .74 .54 .78 .51 .81ZM.43 .25C.4 .23 .4 .18 .43 .15C.46 .13 .51 .13 .54 .15C.57 .18 .57 .23 .54 .25C.51 .28 .46 .28 .43 .25ZM.76 .81C.73 .84 .68 .84 .65 .81C.62 .78 .62 .74 .65 .71C.68 .69 .73 .69 .76 .71C.79 .74 .79 .78 .76 .81Z',
  palette_a: 'M.62 0C.62 .03 .62 .05 .62 .05L.56 .15L.5 .26C.38 .41 .19 .56 .03 .64L0 .64C.03 .77 .06 .87 0 1C.28 .87 .59 .79 .59 .79C.88 .72 1 .49 .88 .26C.88 .26 .81 .18 .62 0Z',
  palette_b: 'M.7 .09C.49 0 .23 .07 .12 .23C0 .39 .09 1 .09 1C.09 1 .79 .7 .88 .54C1 .38 .93 .18 .7 .09Z',
  palette_c: 'M.85 .1C.85 .1 .2 .63 0 .82L.53 1C.67 .78 .95 .13 .95 .13C1 .02 .95 0 .85 .1Z',
  palette_d: 'M1 .58L.08 0C0 .19 0 .48 0 .61C.11 .61 .22 .65 .36 .71C.47 .77 .56 .87 .61 1C.72 .9 .89 .74 1 .58Z',
  clipboard: 'M.51 .07L.48 .07C.47 .07 .46 .06 .46 .05C.46 .02 .43 0 .4 0L.39 0C.36 0 .33 .02 .33 .05C.33 .06 .32 .07 .31 .07L.28 .07C.25 .07 .23 .09 .23 .11C.23 .14 .25 .16 .28 .16L.51 .16C.54 .16 .56 .14 .56 .11C.56 .09 .54 .07 .51 .07ZM.39 .07C.38 .07 .37 .06 .37 .05C.37 .04 .38 .03 .39 .03C.41 .03 .42 .04 .42 .05C.42 .06 .41 .07 .39 .07ZM.51 .9L.07 .9C.03 .9 0 .87 0 .84L0 .16C0 .13 .03 .1 .05 .1L.2 .1L.19 .11C.19 .15 .23 .19 .28 .19L.51 .19C.56 .19 .6 .15 .6 .11C.6 .11 .6 .1 .59 .1L.7 .1C.76 .1 .79 .13 .79 .16L.79 .59C.78 .59 .78 .59 .77 .59C.75 .59 .72 .59 .7 .59L.7 .26L.09 .26L.09 .83L.49 .83C.49 .85 .5 .88 .51 .9ZM.77 .63C.65 .63 .54 .71 .54 .81C.54 .92 .65 1 .77 1C.9 1 1 .92 1 .81C1 .71 .9 .63 .77 .63ZM.92 .78L.78 .9C.77 .91 .76 .91 .75 .91C.75 .91 .74 .91 .73 .91L.64 .85C.63 .84 .63 .82 .64 .81C.65 .8 .67 .8 .69 .81L.75 .85L.87 .74C.88 .73 .9 .73 .92 .74C.93 .75 .93 .76 .92 .78ZM.61 .5L.33 .5C.32 .5 .32 .49 .32 .49L.32 .47C.32 .46 .32 .46 .33 .46L.61 .46C.62 .46 .63 .46 .63 .47L.63 .49C.63 .49 .62 .5 .61 .5ZM.61 .37L.33 .37C.32 .37 .32 .37 .32 .36L.32 .34C.32 .34 .32 .33 .33 .33L.61 .33C.62 .33 .63 .34 .63 .34L.63 .36C.63 .37 .62 .37 .61 .37ZM.53 .61L.33 .61C.32 .61 .32 .61 .32 .6L.32 .59C.32 .58 .32 .57 .33 .57L.53 .57C.54 .57 .54 .58 .54 .59L.54 .6C.54 .61 .54 .61 .53 .61ZM.49 .74L.33 .74C.32 .74 .32 .74 .32 .73L.32 .71C.32 .71 .32 .7 .33 .7L.49 .7C.5 .7 .51 .71 .51 .71L.51 .73C.51 .74 .5 .74 .49 .74ZM.28 .32L.22 .38L.21 .38L.2 .38L.16 .35C.15 .35 .15 .34 .15 .33C.16 .33 .17 .33 .18 .33L.2 .35L.26 .3C.26 .3 .27 .3 .28 .3C.29 .31 .29 .31 .28 .32ZM.28 .45L.22 .51L.21 .51L.2 .51L.16 .48C.15 .48 .15 .47 .15 .46C.16 .46 .17 .46 .18 .46L.2 .48L.26 .43C.26 .42 .27 .42 .28 .43C.29 .43 .29 .44 .28 .45ZM.28 .57L.22 .63L.21 .63L.2 .63L.16 .6C.15 .6 .15 .59 .15 .58C.16 .58 .17 .58 .18 .58L.2 .6L.26 .55C.26 .54 .27 .54 .28 .55C.29 .55 .29 .56 .28 .57ZM.28 .69L.22 .75L.21 .75L.2 .75L.16 .72C.15 .72 .15 .71 .15 .71C.16 .7 .17 .7 .18 .7L.2 .72L.26 .67C.26 .67 .27 .67 .28 .67C.29 .68 .29 .68 .28 .69Z',
  ribbonTailA: 'M1 .57L0 1L.12 0Z',
  ribbonTailB: 'M1 1L0 .97L.02 .78C.06 .52 .11 .26 .13 0Z',
  ribbonLeg: 'M.86 0L.62 0L0 1L.5 1L1 .2Z',
  bar1: 'M0 .41L0 .94C0 .97 .05 1 .11 1L.89 1C.95 1 1 .97 1 .94L1 0L.43 .3C.31 .36 .16 .4 0 .41Z',
  bar2: 'M0 0L0 .94C0 .97 .05 1 .12 1L.89 1C.95 1 1 .97 1 .94L1 .39C.82 .39 .66 .35 .53 .28L0 0Z',
  bar3: 'M0 .44L0 .95C0 .98 .05 1 .11 1L.89 1C.95 1 1 .98 1 .95L1 0L.11 .39C.08 .41 .04 .42 0 .44Z',
  bar4: 'M.92 0L0 .26L0 .97C0 .99 .05 1 .11 1L.88 1C.95 1 1 .99 1 .97L1 .02C.97 .01 .95 .01 .94 0L.92 0Z',
  trendArrow: 'M.99 .02C.99 .01 .98 0 .96 0C.9 0 .85 .01 .79 .01C.78 .01 .77 .01 .76 .03C.76 .03 .76 .04 .75 .04C.75 .07 .76 .09 .76 .1L.78 .13C.79 .15 .8 .16 .81 .18L.46 .76L.3 .5C.29 .48 .28 .47 .27 .47C.25 .47 .24 .48 .23 .5L.02 .84C0 .88 0 .93 .02 .96L.03 .98C.04 .99 .05 1 .06 1C.08 1 .09 .99 .1 .98L.27 .7L.42 .96C.43 .98 .44 .99 .46 .99C.47 .99 .48 .98 .49 .96L.89 .31L.93 .38C.94 .39 .95 .41 .96 .41L.97 .4C.98 .4 .98 .39 .98 .39C.99 .37 .99 .36 .99 .34C.99 .28 .99 .22 1 .15L1 .06C1 .04 1 .03 .99 .02Z',
  house: 'M.96 .59C.95 .59 .95 .59 .94 .59L.93 .59L.5 .14L.07 .59L.06 .59L.05 .59L.01 .53C0 .51 0 .5 .01 .49L.46 .03C.48 0 .52 0 .54 .03L.7 .19L.7 .03C.7 .01 .7 0 .72 0L.84 0C.85 0 .86 .01 .86 .03L.86 .35L.99 .49C1 .5 1 .51 .99 .53L.96 .59ZM.86 .95C.86 .99 .84 1 .82 1L.58 1L.58 .7L.43 .7L.43 1L.19 1C.16 1 .15 .99 .15 .95L.15 .57L.5 .2L.85 .57L.86 .57L.86 .95Z',
  targetRing: 'M.95 .31L.87 .27L.8 .33C.83 .39 .84 .44 .84 .5C.84 .69 .69 .84 .5 .84C.31 .84 .16 .69 .16 .5C.16 .31 .31 .16 .5 .16C.56 .16 .62 .17 .67 .2L.73 .14L.7 .05L.69 .04C.63 .01 .57 0 .5 0C.22 0 0 .22 0 .5C0 .78 .22 1 .5 1C.78 1 1 .78 1 .5C1 .43 .99 .37 .96 .31L.95 .31Z',
  targetArrow: 'M.86 .14C.84 .12 .81 .12 .78 .14L.74 .18L.66 0L.4 .27L.47 .45L.02 .9C0 .92 0 .96 .02 .98C.04 1 .08 1 .1 .98L.56 .53L.73 .6L1 .34L.82 .26L.86 .22C.88 .19 .88 .16 .86 .14Z',
  targetDot: 'M.67 .61C.6 .68 .48 .68 .4 .61C.33 .54 .33 .41 .4 .34L.71 .05C.65 .01 .57 0 .5 0C.22 0 0 .22 0 .5C0 .77 .22 1 .5 1C.77 1 1 .77 1 .5C1 .43 .99 .37 .96 .32L.67 .61Z',
  pole: 'M.41 0C.41 0 .41 0 .53 0C.76 0 1 .01 1 .01C1 .01 1 .01 1 1L.98 1L.01 1L0 1C0 1 0 1 0 .01C0 .01 .18 0 .41 0Z',
  tagBanner: 'M.94 0L.09 0C.04 0 0 .23 0 .5C0 .78 .04 1 .09 1L.94 1L1 .51L.94 0ZM.09 .88C.05 .88 .02 .71 .02 .5C.02 .28 .05 .11 .09 .11C.12 .11 .15 .28 .15 .5C.15 .71 .12 .88 .09 .88Z',
  baseShadow: 'M1 .5C1 .77 .77 1 .5 1C.22 1 0 .77 0 .5C0 .22 .22 0 .5 0C.77 0 1 .22 1 .5',
  baseTop: 'M1 .5C1 .78 .77 .99 .5 .99C.22 .99 0 .78 0 .5C0 .24 .22 0 .5 0C.77 0 1 .24 1 .5',
  thumb: 'M.92 .37C.84 .33 .55 .37 .55 .37C.6 .35 .68 .22 .68 .19C.69 .11 .66 0 .59 0C.51 0 .5 .09 .49 .12C.47 .2 .37 .28 .3 .31C.2 .35 .13 .39 0 .41L0 .42L0 .87C.07 .87 .14 .88 .21 .91C.34 .96 .45 1 .61 1C.85 1 .82 .87 .73 .86C.92 .87 .94 .75 .82 .71C1 .71 .98 .54 .86 .53C.99 .54 1 .41 .92 .37Z',
  trophy: 'M.57 .81C.67 .82 .74 .86 .74 .9C.74 .96 .64 1 .49 1C.36 1 .25 .96 .25 .9C.25 .86 .33 .82 .43 .81L.43 .74C.43 .66 .37 .62 .27 .55C.15 .47 0 .37 0 .14C0 .11 .01 .1 .04 .1L.23 .1C.26 .05 .33 0 .49 0C.66 0 .74 .05 .77 .1L.96 .1C.98 .1 1 .11 1 .14C1 .37 .85 .47 .72 .55C.63 .62 .57 .66 .57 .74L.57 .81ZM.28 .46C.25 .39 .22 .3 .22 .17L.07 .17C.09 .32 .19 .39 .28 .46ZM.29 .15C.29 .17 .35 .23 .49 .23C.65 .23 .71 .17 .71 .15C.71 .13 .65 .07 .49 .07C.35 .07 .29 .13 .29 .15ZM.92 .17L.77 .17C.77 .3 .75 .39 .72 .46C.81 .39 .91 .32 .92 .17Z',
  cart: 'M.81 .79C.76 .79 .73 .82 .71 .87L.39 .87C.37 .84 .35 .82 .34 .81L.39 .68L.9 .68L1 .23L.18 .11L.11 .08L.11 .06C.11 .03 .1 0 .06 0C.02 0 0 .03 0 .06C0 .1 .02 .13 .06 .13L.15 .18L.27 .68L.31 .68L.27 .79C.21 .79 .16 .84 .16 .89C.16 .95 .21 1 .27 1C.32 1 .35 .97 .37 .94L.71 .94C.73 .97 .76 1 .81 1C.87 1 .92 .95 .92 .89C.92 .84 .87 .79 .81 .79Z',
  person: 'M.51 .15C.62 .15 .72 .12 .72 .07C.72 .03 .62 0 .51 0C.39 0 .29 .03 .29 .07C.29 .12 .39 .15 .51 .15ZM1 .25C1 .21 .91 .18 .8 .18L.19 .18C.1 .18 0 .21 0 .25L0 .53C0 .54 .05 .55 .08 .55C.13 .55 .17 .54 .17 .53L.17 .27L.2 .27L.2 .96C.2 .98 .27 1 .34 1C.41 1 .46 .98 .46 .96L.46 .58L.53 .58L.53 .95C.53 .98 .6 1 .67 1C.73 1 .79 .98 .79 .95L.79 .27L.83 .27L.83 .53C.83 .54 .86 .55 .91 .55C.96 .55 1 .54 1 .53L1 .25Z',
  facebook: 'M.82 .19C.68 .19 .66 .24 .66 .28L.66 .35L1 .35L.98 .55L.66 .55L.66 1L.25 1L.25 .55L0 .55L0 .35L.25 .35L.25 .19C.25 .07 .39 0 .66 0L1 0L1 .19L.82 .19Z',
  twitter: 'M.89 .38C.94 .36 .98 .34 1 .28C.97 .3 .91 .32 .88 .3L.88 .28C.84 .17 .75 .08 .66 .08C.67 .08 .67 .08 .69 .08C.69 .06 .75 .06 .75 .02C.73 0 .66 .04 .66 .04C.67 .04 .7 .02 .7 0C.67 0 .64 .02 .62 .04C.62 .04 .64 .02 .64 0C.56 .06 .52 .19 .48 .3C.45 .26 .42 .25 .41 .23C.34 .19 .27 .15 .16 .09C.14 .13 .17 .21 .23 .25C.22 .25 .19 .25 .17 .25C.19 .3 .2 .34 .28 .38C.25 .38 .23 .38 .22 .4C.23 .43 .27 .49 .34 .47C.27 .51 .31 .58 .38 .58C.27 .72 .09 .7 0 .58C.25 1 .8 .83 .88 .43C.94 .43 .98 .42 1 .38C.97 .4 .91 .38 .89 .38Z',
};

/* ------------------------------------------------------------------ helpers */

// PowerPoint's default text insets, in points: [left, right, bottom, top].
const INSETS = [7.2, 7.2, 3.6, 3.6];

/** Text frame matching PowerPoint defaults: top-anchored, standard insets. */
function textBox(s, body, x, y, w, h, opts) {
  s.addText(body, Object.assign({
    x: x, y: y, w: w, h: h,
    margin: INSETS, valign: 'top', align: 'left', isTextBox: true,
    fontFace: TEXT, fontSize: 18, color: INK,
  }, opts || {}));
}

/** Display heading set in Staatliches. */
function head(s, body, x, y, w, h, size, opts) {
  textBox(s, body, x, y, w, h, Object.assign({ fontFace: DISPLAY, fontSize: size }, opts));
}

/** Plain Open Sans line (subheadings, captions, quotes). */
function note(s, body, x, y, w, h, size, opts) {
  textBox(s, body, x, y, w, h, Object.assign({ fontSize: size }, opts));
}

/** Grey 12pt lorem paragraph, 150% leading — the deck's standard body copy. */
function body(s, copy, x, y, w, h, opts) {
  textBox(s, copy, x, y, w, h, Object.assign({
    fontSize: 12, color: MUTED, charSpacing: 1, lineSpacingMultiple: 1.5,
  }, opts));
}

/** Big bold Open Sans number (prices, counters, list indices). */
function stat(s, body, x, y, w, h, size, opts) {
  textBox(s, body, x, y, w, h, Object.assign({ fontSize: size, bold: true }, opts));
}

/** Bulleted feature list, 200% leading. */
function bullets(s, items, x, y, w, h, opts) {
  const o = Object.assign({ color: INK, lineSpacingMultiple: 2 }, opts);
  textBox(s, items.map(function (t) {
    return { text: t, options: { breakLine: true, bullet: { indent: 13.5 } } };
  }), x, y, w, h, Object.assign({ fontSize: 12, charSpacing: 1 }, o));
}

/** Thin grey rule (the deck's signature ornament stroke). */
function rule(s, x, y, w, h, opts) {
  s.addShape('line', Object.assign({ x: x, y: y, w: w, h: h, line: HAIRLINE }, opts));
}

function box(s, x, y, w, h, color, opts) {
  s.addShape('rect', Object.assign({ x: x, y: y, w: w, h: h, fill: { color: color }, line: NO_LINE }, opts));
}

function frame(s, x, y, w, h, color, opts) {
  s.addShape('rect', Object.assign({ x: x, y: y, w: w, h: h, fill: { type: 'none' }, line: { color: color, width: 2.25 } }, opts));
}

function dot(s, x, y, w, h, color, opts) {
  s.addShape('ellipse', Object.assign({ x: x, y: y, w: w, h: h, fill: { color: color }, line: NO_LINE }, opts));
}

function tri(s, x, y, w, h, color, opts) {
  s.addShape('triangle', Object.assign({ x: x, y: y, w: w, h: h, fill: { color: color }, line: NO_LINE }, opts));
}

function para(s, x, y, w, h, color, adj, opts) {
  s.addShape('parallelogram', Object.assign({
    x: x, y: y, w: w, h: h, fill: { color: color }, line: NO_LINE, rectRadius: adj * Math.min(w, h),
  }, opts));
}

/** Stand-in for a photograph in the reference deck. */
function img(s, x, y, w, h) {
  s.addShape('rect', { x: x, y: y, w: w, h: h, fill: { color: PLACEHOLDER_FILL }, line: NO_LINE });
  s.addText('[image]', {
    x: x, y: y, w: w, h: h, margin: 0, align: 'center', valign: 'middle',
    fontFace: TEXT, fontSize: 14, color: PLACEHOLDER_TEXT,
  });
}

/**
 * Corner ornament: a hairline plus a small rotated caption, one of the 19
 * placements catalogued in CORNERS.
 */
function corner(s, id, label) {
  const c = CORNERS[id];
  rule(s, c.line[0], c.line[1], c.line[2], c.line[3]);
  note(s, label === undefined ? T.orn : label, c.text[0], c.text[1], c.text[2], c.text[3], 18,
    { align: c.align || 'left', rotate: c.rotate || 0, flipH: !!c.flipH, flipV: !!c.flipV });
}

/**
 * Draw a normalised outline from PATHS. Coordinates in the path string run
 * 0..1 across the shape's own bounding box and are scaled to w/h here;
 * custGeom points are shape-local, so x/y only position the bounding box.
 */
function icon(s, d, x, y, w, h, color, opts) {
  const pts = [];
  const re = /([MLCZ])([-0-9. ]*)/g;
  let m;
  while ((m = re.exec(d)) !== null) {
    if (m[1] === 'Z') { pts.push({ close: true }); continue; }
    const v = m[2].trim().split(' ').map(Number);
    if (m[1] === 'M') pts.push({ x: v[0] * w, y: v[1] * h, moveTo: true });
    else if (m[1] === 'L') pts.push({ x: v[0] * w, y: v[1] * h });
    else pts.push({
      x: v[4] * w, y: v[5] * h,
      curve: { type: 'cubic', x1: v[0] * w, y1: v[1] * h, x2: v[2] * w, y2: v[3] * h },
    });
  }
  s.addShape('custGeom', Object.assign({
    x: x, y: y, w: w, h: h, points: pts, fill: { color: color }, line: NO_LINE,
  }, opts));
}

/* ------------------------------------------------------------------- slides */
const WHITE_SLIDES = [3, 13, 21, 30, 40, 49, 50];   // slides whose background is plain white

const SLIDES = [
  /* 1 — ORNAMENT */
  function (s) {
    img(s, 0, 3.75, 7.046, 3.75);
    img(s, 6.287, 0, 7.046, 4.244);
    head(s, 'ORNAMENT', 1.667, 2.447, 10, 1.192, 80, { align: 'center' });
    note(s, 'Presentation Template', 1.667, 3.861, 10, .493, 20, { color: '000000', charSpacing: 3, align: 'center' });
    corner(s, 'C1', '2019');
    rule(s, 12.654, 6.023, 0, 1.477);
    note(s, T.tag, 10.609, 3.378, 4.09, .404, 18, { rotate: 270 });
  },

  /* 2 — Simplicity */
  function (s) {
    note(s, 'Leonardo', 1.656, 5.952, 1.984, .505, 24);
    rule(s, 0, 6.639, 3.246, 0);
    img(s, 3.967, 3.75, 8.396, 3.75);
    head(s, 'Simplicity', 1.683, 1.117, 4.317, 1.212, 66);
    note(s, 'is the Ultimate', 1.683, 2.375, 3.792, .707, 36);
    head(s, 'Sophistication', 1.683, 3.164, 5.407, 1.212, 66);
    corner(s, 'C2', T.orn);
  },

  /* 3 — Be Faithful */
  function (s) {
    img(s, 2.874, 0, 10.459, 7.5);
    head(s, 'Be Faithful', 1.093, 2.538, 4.317, 1.212, 66);
    note(s, 'to your own taste', 1.093, 3.764, 5.104, .707, 36);
    note(s, 'Billy Baldwin', 1.093, 5.091, 2.609, .505, 24);
    rule(s, 0, 5.754, 3.246, 0);
    corner(s, 'C3', T.orn);
  },

  /* 4 — Furniture */
  function (s) {
    img(s, 1.149, 1.82, 10.459, 5.68);
    note(s, 'in a modern loft, you can’t just fill a space with', 2.104, .738, 3.792, 2.524, 36);
    head(s, 'Furniture', 8.404, 3.144, 4.317, 1.212, 66, { align: 'right' });
    note(s, 'Nate Berkus', 10.126, 4.493, 2.596, .505, 24, { align: 'right' });
    rule(s, 10.738, 5.229, 2.596, 0);
    corner(s, 'C4', T.orn);
  },

  /* 5 — TABLE of CONTENT */
  function (s) {
    img(s, .951, 1.115, 5.869, 4.873);
    head(s, 'TABLE of CONTENT', 1.647, 5.108, 3.24, 2.121, 60, { color: HEADLINE });
    body(s, T.p4, 7.678, 2.877, 4.322, .669);
    head(s, 'ABOUT US', 7.678, 2.049, 4.322, .841, 44);
    stat(s, '01', 6.344, 1.914, 1.268, 1.111, 60, { align: 'center' });
    body(s, T.p4, 7.678, 4.988, 3.978, .669);
    head(s, 'PRODUCT & SERVICES', 7.678, 4.16, 4.861, .841, 44);
    stat(s, '02', 6.344, 4.025, 1.268, 1.111, 60, { align: 'center' });
    rule(s, 0, 6.139, 4.361, 0);
    corner(s, 'C3', T.orn);
  },

  /* 6 — TABLE of CONTENT */
  function (s) {
    img(s, 6.508, 1.115, 5.869, 4.873);
    head(s, 'TABLE of CONTENT', 8.434, 5.112, 3.24, 2.121, 60, { color: HEADLINE, align: 'right' });
    body(s, T.p4, 1.366, 2.475, 4.322, .669, { align: 'right' });
    head(s, 'SWOT', 1.366, 1.647, 4.322, .841, 44, { align: 'right' });
    stat(s, '03', 5.787, 1.512, 1.268, 1.111, 60, { align: 'right' });
    body(s, T.p4, .464, 4.718, 5.222, .669, { align: 'right' });
    head(s, 'MARKET & MANAGEMENT', .464, 3.889, 5.222, .841, 44, { align: 'right' });
    stat(s, '04', 5.787, 3.755, 1.268, 1.111, 60, { align: 'right' });
    rule(s, 8.967, 6.139, 4.361, 0);
    corner(s, 'C5', T.orn);
  },

  /* 7 — TABLE of CONTENT */
  function (s) {
    img(s, .963, 2.546, 7.012, 4.213);
    head(s, 'TABLE of CONTENT', 8.991, .518, 3.24, 2.121, 60, { color: HEADLINE, align: 'right' });
    rule(s, 9.475, 1.566, 3.869, 0);
    body(s, T.p4, 7.776, 3.897, 4.322, .669);
    head(s, 'TEAM', 7.776, 3.068, 4.322, .841, 44);
    stat(s, '05', 6.442, 2.934, 1.268, 1.111, 60, { align: 'center' });
    body(s, T.p4, 7.776, 6.008, 3.978, .669);
    head(s, 'OUR GALLERY', 7.776, 5.18, 4.861, .841, 44);
    stat(s, '06', 6.442, 5.045, 1.268, 1.111, 60, { align: 'center' });
    corner(s, 'C4', T.orn);
  },

  /* 8 — Welcome Message ! */
  function (s) {
    img(s, 5.215, 0, 8.118, 5.462);
    head(s, 'Welcome Message !', 1.8, 1.24, 4.317, 2.121, 60, { align: 'right' });
    body(s, T.p1, .871, 4.56, 3.606, 2.189);
    rule(s, 0, 3.606, 1.8, 0);
    note(s, T.sub, 2.027, 3.408, 4.09, .404, 18, { align: 'right' });
    corner(s, 'C6', T.orn);
  },

  /* 9 — WHO WE are? */
  function (s) {
    img(s, 2.541, 0, 5.88, 7.5);
    head(s, 'WHO WE are?', .887, 2.286, 3.165, 2.121, 60);
    body(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae', 8.976, 3.683, 3.606, 1.886);
    note(s, T.sub, .887, 4.549, 2.622, .707, 18);
    rule(s, -.018, 5.466, 3.053, 0);
    corner(s, 'C2', T.orn);
  },

  /* 10 — About us */
  function (s) {
    img(s, 1, 0, 5.667, 5.035);
    head(s, 'About us', 1.386, 5.158, 4.895, 1.111, 60, { align: 'center' });
    note(s, T.sub, 1.198, 6.339, 5.271, .404, 18, { align: 'center' });
    body(s, T.p1, 8.09, 2.798, 3.954, 1.886, { align: 'center' });
    rule(s, 10.067, -.018, 0, 2.193);
    rule(s, 10.067, 5.307, 0, 2.193);
    corner(s, 'C7', T.orn);
  },

  /* 11 — Our  vision */
  function (s) {
    img(s, 2.833, 0, 7.667, 7.5);
    head(s, 'Our  vision', 4.219, 2.958, 4.895, 1.111, 60, { align: 'center' });
    note(s, T.sub, 4.031, 4.139, 5.271, .404, 18, { align: 'center' });
    body(s, T.p3, .82, 1.304, 3.549, 1.583);
    head(s, 'Vision ONE', .82, .492, 3.549, .707, 36);
    body(s, T.p3, 8.947, 5.355, 3.584, 1.583, { align: 'right' });
    head(s, 'Vision TWO', 8.947, 4.542, 3.584, .707, 36, { align: 'right' });
    corner(s, 'C8', T.orn);
    corner(s, 'C3', '2019');
  },

  /* 12 — Our  MISSion */
  function (s) {
    img(s, 2.833, 0, 7.667, 7.5);
    head(s, 'Our  MISSion', 4.219, 2.958, 4.895, 1.111, 60, { align: 'center' });
    note(s, T.sub, 4.031, 4.139, 5.271, .404, 18, { align: 'center' });
    body(s, T.p3, .82, 1.304, 3.399, 1.583);
    head(s, 'MISsion ONE', .82, .492, 3.399, .707, 36);
    body(s, T.p3, 9.114, 5.355, 3.417, 1.583, { align: 'right' });
    head(s, 'MISSion TWO', 9.114, 4.542, 3.417, .707, 36, { align: 'right' });
    corner(s, 'C3', T.orn);
    corner(s, 'C8', '2019');
  },

  /* 13 — Break slide! */
  function (s) {
    img(s, 0, 0, 7.869, 6.59);
    head(s, 'Break slide!', 7.027, 1.884, 4.514, 2.794, 80);
    note(s, T.tag, 7.027, 4.843, 4.514, .404, 18);
    rule(s, 7.027, 5.511, 6.306, 0);
    corner(s, 'C7', T.orn);
  },

  /* 14 — GOOD MATERIALS */
  function (s) {
    img(s, 1.461, .496, 5.066, 6.508);
    body(s, T.p7, 8.044, 1.509, 3.927, .978);
    head(s, 'GOOD MATERIALS', 8.044, .762, 3.549, .707, 36);
    head(s, 'PRODUCT FEATURES', .562, 4.615, 3.914, 2.121, 60);
    note(s, T.sub, 4.389, 3.089, 5.066, .404, 18, { rotate: 270 });
    body(s, T.p7, 8.044, 3.592, 3.927, .978);
    head(s, 'ECHO FRIENDLY', 8.044, 2.845, 3.549, .707, 36);
    body(s, T.p7, 8.044, 5.675, 3.923, .978);
    head(s, 'AMAZING DESIGN', 8.044, 4.928, 3.549, .707, 36);
    corner(s, 'C4', T.orn);
    rule(s, 6.918, 6.056, 0, 1.477);
  },

  /* 15 — LIVING ROOM */
  function (s) {
    img(s, 8.661, 2.93, 4.454, 4.016);
    img(s, .218, 2.93, 4.454, 4.016);
    img(s, 4.45, 2.365, 4.433, 3.998);
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '750' }], 10.044, 1.916, 1.909, .909, 48, { align: 'right' });
    head(s, 'LIVING ROOM', 10.011, 3.132, 2.943, .707, 36, { align: 'right' });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '640' }], .342, 1.913, 1.655, .909, 48, { align: 'right' });
    head(s, 'BATH ROOM', .342, 3.129, 2.943, .707, 36);
    head(s, 'BED ROOM', 5.195, 6.371, 2.943, .707, 36, { align: 'center' });
    head(s, 'OUR PRODUCTS', 4.054, .393, 5.225, 1.111, 60, { align: 'center' });
    note(s, T.sub, 4.57, 1.467, 4.194, .404, 18, { align: 'center' });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '850' }], 5.67, 1.929, 1.863, 1.01, 54, { align: 'right' });
    rule(s, 0, .882, 3.887, 0);
    rule(s, 9.446, .882, 3.887, 0);
    rule(s, 1.874, 2.365, .46, .532, { flipH: true });
    head(s, 'Set', 2.288, 2.454, .813, .438, 20);
    rule(s, 7.366, 2.54, .46, .532, { flipH: true });
    head(s, 'Set', 7.779, 2.629, .813, .438, 20);
    rule(s, 11.952, 2.366, .46, .532, { flipH: true });
    head(s, 'Set', 12.366, 2.456, .749, .438, 20);
  },

  /* 16 — BATH ROOM */
  function (s) {
    box(s, 9.849, 0, 3.018, 7.5, 'D9D9D9');
    head(s, 'BATH ROOM', .886, 4.628, 3.981, 1.212, 66);
    note(s, T.sub, .886, 5.84, 4.194, .404, 18);
    body(s, T.p3, 5.584, 1.059, 3.549, 1.583);
    bullets(s, [T.p2, T.p2, T.p2], 9.849, 4.123, 3.018, 1.919, { lineSpacingMultiple: 1.5 });
    corner(s, 'C2', T.orn);
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '640' }], 9.62, 1.602, 2.236, 1.01, 54, { align: 'right' });
    rule(s, 11.803, 2.259, .46, .532, { flipH: true });
    head(s, 'Set', 12.217, 2.348, .813, .438, 20);
    img(s, .667, .517, 4.2, 3.381);
    img(s, 5.259, 3.543, 4.2, 3.382);
  },

  /* 17 — LIVING ROOM */
  function (s) {
    bullets(s, [T.p2, T.p2, T.p2], 3.784, 4.767, 3.018, 1.919, { lineSpacingMultiple: 1.5 });
    body(s, T.p3, 8.131, 2.494, 4.148, 1.28);
    head(s, 'LIVING ROOM', 8.131, .643, 4.574, 1.212, 66);
    note(s, T.sub, 8.131, 1.855, 4.194, .404, 18);
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '750' }], .743, 5.12, 2.087, 1.01, 54, { align: 'right' });
    rule(s, 2.64, 5.748, .46, .532, { flipH: true });
    head(s, 'Set', 3.053, 5.838, .813, .438, 20);
    corner(s, 'C4', T.orn);
    img(s, 1.023, 0, 5.765, 4.115);
    img(s, 7.705, 4.262, 5, 3.238);
  },

  /* 18 — BED ROOM */
  function (s) {
    box(s, .522, 0, 3.018, 7.5, 'D9D9D9');
    head(s, 'BED ROOM', 8.884, 4.628, 3.981, 1.212, 66, { align: 'right' });
    note(s, T.sub, 8.671, 5.84, 4.194, .404, 18, { align: 'right' });
    body(s, T.p3, 4.403, 1.059, 3.549, 1.583, { align: 'right' });
    bullets(s, [T.p2, T.p2, T.p2], .522, 4.123, 3.018, 1.919, { lineSpacingMultiple: 1.5 });
    corner(s, 'C7', T.orn);
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '850' }], .241, 1.709, 2.486, 1.01, 54, { align: 'right' });
    rule(s, 2.424, 2.366, .46, .532, { flipH: true });
    head(s, 'Set', 2.838, 2.456, .813, .438, 20);
    img(s, 8.477, 0, 4.2, 3.898);
    img(s, 4.471, 3.544, 4.2, 3.956);
  },

  /* 19 — SPECIAL PRICE */
  function (s) {
    img(s, 2.79, 0, 7.754, 7.5);
    head(s, 'SPECIAL PRICE', 1.562, 1.163, 3.914, 2.121, 60);
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '640' }], 9.23, 1.611, 2.437, 1.01, 54, { align: 'right' });
    rule(s, 11.607, 2.354, .46, .532, { flipH: true });
    head(s, 'Set', 12.021, 2.443, .813, .438, 20);
    body(s, 'Cras vel mauris et tortor tincidunt pulvinar. Curabitur ac ullamcorper nisi. Nullam vel venenatis purus. In lobortis ipsum eget ante rhoncus dignissim. Sed sed venenatis lectus. Nulla tempor efficitur arcu et efficitur. ', .681, 3.657, 3.106, 2.492);
    head(s, 'NIGHT-LAMP SET', 9.503, 3.165, 3.549, .707, 36);
    rule(s, 1.374, 0, 0, 3.058);
    corner(s, 'C6', T.orn);
  },

  /* 20 — BIG SALE */
  function (s) {
    img(s, 1.045, .652, 11.243, 6.197);
    head(s, 'BIG SALE', .545, .918, 2.182, 2.121, 60);
    corner(s, 'C9', T.orn);
    corner(s, 'C8', '2019');
    box(s, 5.745, 5.291, 2.758, 1.163, 'D9D9D9', { transparency: 23.9 });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '90' }], 5.907, 5.47, 1.207, .707, 36, { align: 'right' });
    head(s, 'MODERN CHAIR', 7.16, 5.487, 1.207, .774, 20);
    box(s, 7.309, .864, 3.055, 1.163, 'D9D9D9', { transparency: 23.9 });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '100' }], 7.16, 1.06, 1.517, .707, 36, { align: 'right' });
    head(s, 'Minimalist Black  lamp', 8.723, 1.059, 1.531, .774, 20);
    box(s, 2.909, 3.015, 3.277, 1.163, 'D9D9D9', { transparency: 23.9 });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '390' }], 2.924, 3.194, 1.505, .707, 36, { align: 'right' });
    head(s, 'MODERN KITCHEN SET', 4.475, 3.211, 1.505, .774, 20);
    box(s, 8.341, 2.689, 2.758, 1.163, 'D9D9D9', { transparency: 23.9 });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '80' }], 8.502, 2.867, 1.207, .774, 40, { align: 'right' });
    head(s, 'OSTRICH PAINTING', 9.755, 2.884, 1.207, .774, 20);
  },

  /* 21 — slide! */
  function (s) {
    img(s, 0, 0, 7.509, 7.5);
    head(s, 'slide!', 6.966, 5.319, 4.514, 1.447, 80);
    rule(s, 7.027, 5.167, 6.306, 0);
    corner(s, 'C3', T.orn);
    note(s, T.tag, -1.676, 1.92, 4.617, .404, 18, { rotate: 270 });
    rule(s, .613, 4.656, 0, 2.844);
    head(s, 'Break', 6.966, 3.261, 4.514, 2.036, 115);
  },

  /* 22 — RESEARCH */
  function (s) {
    body(s, T.p5, 2.672, 3.364, 3.549, 1.275);
    head(s, 'RESEARCH', 2.672, 2.552, 3.549, .707, 36);
    body(s, T.p5, 7.14, 3.364, 3.549, 1.275);
    head(s, 'ECHO FRIENDLY', 7.14, 2.552, 3.549, .707, 36);
    head(s, 'Our  SERVICES', 1.119, 5.049, 3.914, 2.121, 60);
    note(s, T.sub, 8.656, 1.102, 3.126, .707, 18, { align: 'right' });
    corner(s, 'C4', T.orn);
    img(s, 2.658, 0, 8.03, 7.5);
  },

  /* 23 — MODERN TRENDS */
  function (s) {
    img(s, 3.494, 4.148, 4.869, 3.352);
    body(s, T.p5, 6.667, 1.761, 3.549, 1.275);
    head(s, 'MODERN TRENDS', 6.667, .949, 3.549, .707, 36);
    body(s, T.p5, 8.83, 5.276, 3.549, 1.275);
    head(s, 'GOOD EXECUTION', 8.83, 4.464, 3.549, .707, 36);
    corner(s, 'C3', T.orn);
    note(s, T.sub, -1.676, 1.92, 4.617, .404, 18, { rotate: 270 });
    head(s, 'Our  SERVICES', 1.18, 5, 3.914, 2.121, 60);
    rule(s, .613, 4.656, 0, 2.844);
    img(s, 1.06, 0, 4.869, 5);
  },

  /* 24 — GOOD MATERIALS */
  function (s) {
    img(s, 0, 1.492, 13.333, 4.328);
    body(s, T.p5, 6.027, 1.54, 3.549, 1.275);
    head(s, 'GOOD MATERIALS', 6.027, .728, 3.549, .707, 36);
    body(s, T.p5, 6.027, 5.074, 3.549, 1.275);
    head(s, 'PROFESSIONAL DESIGNER', 6.027, 3.607, 3.549, 1.313, 36);
    head(s, 'Our  SERVICES', .945, .532, 3.914, 2.121, 60);
    note(s, T.sub, 1.126, 5.984, 4.458, .404, 18);
    corner(s, 'C2', T.orn);
    rule(s, 0, 6.175, .945, 0);
  },

  /* 25 — STRATEGY ONE */
  function (s) {
    body(s, T.p6, .995, 1.1, 4.464, .972);
    head(s, 'STRATEGY ONE', .995, .338, 3.549, .707, 36);
    body(s, T.p6, 5.549, 1.1, 4.464, .972);
    head(s, 'STRATEGY TWO', 5.549, .338, 3.549, .707, 36);
    body(s, T.p6, .995, 5.988, 4.464, .972);
    head(s, 'STRATEGY THREE', .995, 5.226, 3.549, .707, 36);
    body(s, T.p6, 5.549, 5.988, 4.464, .972);
    head(s, 'STRATEGY FOUR', 5.549, 5.226, 3.549, .707, 36);
    note(s, T.sub, 8.985, 4.783, 3.914, .707, 18, { align: 'right' });
    head(s, 'MARKET STRATEGY', 8.985, 2.435, 3.914, 2.121, 60, { align: 'right' });
    corner(s, 'C10', T.orn);
    rule(s, 9.632, 4.582, 3.702, 0);
    img(s, 0, 2.26, 8.967, 2.882);
  },

  /* 26 — SWOT ANALYSIS */
  function (s) {
    img(s, 1.254, 0, 3.678, 7.5);
    note(s, T.sub, 7.761, 6.178, 5.137, .404, 18);
    head(s, 'SWOT ANALYSIS', 7.761, 5.067, 5.137, 1.111, 60);
    head(s, 'S', 4.146, 0, 1.773, 7.052, 413);
    body(s, T.p1, 7.761, 2.182, 3.606, 2.189);
    head(s, 'STRENGTH', 7.761, 1.457, 3.549, .64, 32);
    corner(s, 'C3', T.orn);
  },

  /* 27 — SWOT ANALYSIS */
  function (s) {
    img(s, 2.415, 0, 4.374, 7.5);
    note(s, T.sub, 5.575, 1.632, 5.137, .404, 18);
    head(s, 'SWOT ANALYSIS', 5.575, .521, 5.137, 1.111, 60);
    head(s, 'W', .493, .528, 1.773, 7.052, 413);
    body(s, T.p1, 7.592, 3.478, 3.606, 2.189);
    head(s, 'WEAKNESS', 7.592, 2.754, 3.549, .64, 32);
    corner(s, 'C11', T.orn);
  },

  /* 28 — SWOT ANALYSIS */
  function (s) {
    img(s, 8.576, 0, 4.091, 7.5);
    note(s, T.sub, -.623, 3.994, 5.137, .404, 18, { rotate: 270 });
    head(s, 'SWOT ANALYSIS', -1.381, 3.64, 5.137, 1.111, 60, { rotate: 270 });
    head(s, 'O', 7.094, .445, 1.773, 7.052, 413);
    body(s, T.p1, 3.066, 3.062, 3.606, 2.189);
    head(s, 'OPPORTUNITIES', 3.066, 2.338, 3.549, .64, 32);
    corner(s, 'C12', T.orn);
  },

  /* 29 — SWOT ANALYSIS */
  function (s) {
    img(s, 0, 3.884, 7.639, 3.616);
    note(s, T.sub, 8.033, 5.807, 5.137, .404, 18);
    head(s, 'SWOT ANALYSIS', 8.033, 4.697, 5.137, 1.111, 60);
    head(s, 'T', 1.292, -.276, 1.773, 7.052, 413);
    body(s, T.p1, 5.28, 1.666, 4.834, 1.583);
    head(s, 'THREATS', 5.28, .942, 3.549, .64, 32);
    corner(s, 'C3', T.orn);
  },

  /* 30 — Break slide! */
  function (s) {
    img(s, 1.045, .652, 11.243, 6.197);
    head(s, 'Break slide!', 1.73, 2.781, 5.874, 1.447, 80, { align: 'center' });
    note(s, T.tag, 2.41, 4.316, 4.514, .404, 18, { align: 'center' });
    rule(s, 1.514, 4.195, 6.306, 0);
    corner(s, 'C9', T.orn);
    corner(s, 'C8', '2019');
  },

  /* 31 — PRICING TABLE */
  function (s) {
    box(s, .672, 0, 4.197, 7.5, INK);
    box(s, 5.396, 0, 4.197, 7.5, INK);
    note(s, T.sub, 9.957, 5.441, 3.148, .707, 18);
    head(s, 'PRICING TABLE', 9.957, 3.255, 3.148, 2.121, 60);
    bullets(s, ['Item Detail 01', 'Item Detail 02', 'Item Detail 03', 'Item Detail 04', 'Item Detail 05', 'Item Detail 06', 'Item Detail 07', 'Item Detail 08'], 1.865, 2.406, 2.003, 3.269, { color: CANVAS });
    note(s, 'Price', 1.395, 1.536, 1.005, .505, 24, { fontFace: 'Staatliches', color: CANVAS, charSpacing: 1, bold: true });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '250' }], 2.222, 1.233, 2.07, 1.01, 54, { color: CANVAS, charSpacing: 3, align: 'center' });
    head(s, 'Plan a', .672, 6.089, 4.197, .707, 36, { color: PAPER, align: 'center' });
    head(s, 'Plan B', 5.399, 6.078, 4.197, .707, 36, { color: PAPER, align: 'center' });
    bullets(s, ['Item Detail 01', 'Item Detail 02', 'Item Detail 03', 'Item Detail 04', 'Item Detail 05', 'Item Detail 06', 'Item Detail 07'], 6.514, 2.406, 2.003, 2.865, { color: CANVAS });
    note(s, 'Price', 6.044, 1.536, 1.005, .505, 24, { fontFace: 'Staatliches', color: CANVAS, charSpacing: 1, bold: true });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '200' }], 6.871, 1.233, 2.07, 1.01, 54, { color: CANVAS, charSpacing: 3, align: 'center' });
    corner(s, 'C2', T.orn);
  },

  /* 32 — Plan a */
  function (s) {
    frame(s, .96, .459, 3.752, 6.623, INK);
    bullets(s, ['Item Detail 01', 'Item Detail 02', 'Item Detail 03', 'Item Detail 04', 'Item Detail 05', 'Item Detail 06', 'Item Detail 07', 'Item Detail 08'], 1.931, 2.225, 2.003, 3.269);
    note(s, 'Price', 1.461, 1.356, 1.005, .505, 24, { fontFace: 'Staatliches', charSpacing: 1, bold: true });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '550' }], 2.288, 1.053, 2.07, 1.01, 54, { charSpacing: 3, align: 'center' });
    head(s, 'Plan a', .96, 5.909, 3.752, .707, 36, { align: 'center' });
    frame(s, 5.088, 1.41, 3.752, 5.672, INK);
    bullets(s, ['Item Detail 01', 'Item Detail 02', 'Item Detail 03', 'Item Detail 04', 'Item Detail 05', 'Item Detail 06'], 6.058, 3.111, 2.003, 2.462);
    note(s, 'Price', 5.589, 2.241, 1.005, .505, 24, { fontFace: 'Staatliches', charSpacing: 1, bold: true });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '400' }], 6.416, 1.938, 2.07, 1.01, 54, { charSpacing: 3, align: 'center' });
    head(s, 'Plan B', 5.088, 5.876, 3.752, .707, 36, { align: 'center' });
    frame(s, 9.146, 2.811, 3.752, 4.254, INK);
    bullets(s, ['Item Detail 01', 'Item Detail 02', 'Item Detail 03', 'Item Detail 04'], 10.116, 4.291, 2.003, 1.654);
    note(s, 'Price', 9.647, 3.421, 1.005, .505, 24, { fontFace: 'Staatliches', charSpacing: 1, bold: true });
    stat(s, [{ text: '$', options: { baseline: 600 } }, { text: '300' }], 10.473, 3.118, 2.07, 1.01, 54, { charSpacing: 3, align: 'center' });
    head(s, 'Plan C', 9.146, 6.138, 3.752, .707, 36, { align: 'center' });
    note(s, T.sub, 5.933, .375, 3.148, .707, 18, { align: 'right' });
    head(s, 'PRICING TABLE', 9.313, .277, 3.148, 2.121, 60);
    rule(s, 9.216, -.015, 0, 2.193);
    corner(s, 'C8', T.orn);
  },

  /* 33 — PROCESS FLOW */
  function (s) {
    box(s, .71, 2.197, 2.164, 2.164, INK);
    box(s, 3.107, 4.806, 2.164, 2.164, '737373');
    box(s, 5.585, 2.197, 2.164, 2.164, INK);
    box(s, 8.074, 4.806, 2.164, 2.164, '737373');
    box(s, 10.459, 2.197, 2.164, 2.164, INK);
    rule(s, 1.772, 4.588, 9.79, 0);
    dot(s, 1.649, 4.466, .246, .246, INK);
    dot(s, 11.439, 4.466, .246, .246, INK);
    head(s, 'PROCESS FLOW', 4.074, .287, 5.184, 1.111, 60, { align: 'center' });
    note(s, T.sub, 3.986, 1.398, 5.362, .404, 18, { align: 'center' });
    icon(s, P.flag, 1.565, 2.517, .66, .791, 'EBECEE');
    icon(s, P.search, 3.867, 5.301, .644, .636, 'EBECEE');
    icon(s, P.hammer, 8.89, 5.301, .569, .702, 'EBECEE');
    icon(s, P.palette, 6.41, 2.681, .561, .612, 'EBECEE', { rotate: 7.82 });
    icon(s, P.palette_a, 6.897, 2.807, .109, .131, 'EBECEE', { rotate: 7.82 });
    icon(s, P.palette_b, 6.741, 2.828, .145, .191, 'EBECEE', { rotate: 7.82 });
    icon(s, P.palette_c, 6.869, 2.503, .205, .308, 'EBECEE', { rotate: 7.82 });
    icon(s, P.palette_d, 6.823, 2.767, .121, .105, 'EBECEE', { rotate: 7.82 });
    icon(s, P.clipboard, 11.202, 2.556, .678, .831, 'EBECEE');
    head(s, 'START', .71, 3.636, 2.164, .572, 28, { color: CANVAS, align: 'center' });
    dot(s, 4.074, 4.466, .246, .246, '737373');
    dot(s, 6.544, 4.462, .246, .246, INK);
    dot(s, 9.052, 4.462, .246, .246, '737373');
    head(s, 'RESEARCH', 3.107, 6.2, 2.164, .572, 28, { color: CANVAS, align: 'center' });
    head(s, 'DESIGN', 5.585, 3.636, 2.164, .572, 28, { color: CANVAS, align: 'center' });
    head(s, 'RELEASE', 10.459, 3.636, 2.164, .572, 28, { color: CANVAS, align: 'center' });
    head(s, 'CREATE', 8.078, 6.2, 2.164, .572, 28, { color: CANVAS, align: 'center' });
    corner(s, 'C13', '2019');
    corner(s, 'C2', T.orn);
  },

  /* 34 — ADVERTISING */
  function (s) {
    icon(s, P.ribbonTailA, 4.728, 5.635, .677, 1.199, '737373', { rotate: 123.03, flipH: true });
    tri(s, 3.793, 1.178, 1.258, .358, '737373');
    icon(s, P.ribbonTailB, .987, 3.89, .775, 1.402, '737373', { rotate: 267.17 });
    para(s, .831, 1.178, 3.594, 4.701, 'D9D9D9', .553);
    box(s, 2.775, 2.83, 1.424, 5.582, 'BFBFBF', { rotate: 90 });
    icon(s, P.ribbonLeg, 3.835, 1.535, 3.202, 5.243, 'A6A6A6', { flipH: true });
    icon(s, P.bar1, 2.362, 3.341, .099, .187, INK);
    icon(s, P.bar2, 2.5, 3.339, .101, .189, INK);
    icon(s, P.bar3, 2.639, 3.304, .1, .225, INK);
    icon(s, P.bar4, 2.777, 3.173, .1, .355, INK);
    icon(s, P.trendArrow, 2.316, 3.003, .623, .378, INK);
    icon(s, P.house, 2.685, 5.353, .661, .525, INK);
    icon(s, P.targetRing, 4.931, 3.911, .547, .55, INK);
    icon(s, P.targetArrow, 5.19, 3.829, .374, .374, INK);
    icon(s, P.targetDot, 5.069, 4.051, .271, .271, INK);
    body(s, T.p8, 7.948, 1.316, 4.315, .978);
    head(s, 'ADVERTISING', 7.948, .591, 3.549, .64, 32);
    body(s, T.p8, 7.948, 3.596, 4.315, .978);
    head(s, 'STRUCTURE', 7.948, 2.872, 3.549, .64, 32);
    body(s, T.p8, 7.948, 5.877, 4.315, .978);
    head(s, 'STRATEGY', 7.948, 5.152, 3.549, .64, 32);
    head(s, 'BUSINESS MODEL', .558, .52, 3.549, 2.121, 60);
    stat(s, '01', 6.711, .456, 1.268, .909, 48, { align: 'center' });
    stat(s, '02', 6.711, 2.737, 1.268, .909, 48, { align: 'center' });
    stat(s, '03', 6.711, 5.023, 1.268, .909, 48, { align: 'center' });
    note(s, T.sub, .675, 6.578, 4.315, .404, 18);
    corner(s, 'C6', T.orn);
  },

  /* 35 — PLAN & DESIGN */
  function (s) {
    img(s, 7.049, 1.253, 4.06, 5.798);
    body(s, T.p9, 1.377, 1.382, 3.944, .978, { align: 'right' });
    head(s, 'PLAN & DESIGN', 1.772, .658, 3.549, .64, 32, { align: 'right' });
    body(s, T.p9, 1.377, 3.663, 3.944, .978, { align: 'right' });
    head(s, 'DESIGN', 1.772, 2.939, 3.549, .64, 32, { align: 'right' });
    body(s, T.p9, 1.377, 5.944, 3.944, .978, { align: 'right' });
    head(s, 'DEVELOPMENT', 1.772, 5.219, 3.549, .64, 32, { align: 'right' });
    stat(s, '01', 5.486, .523, 1.268, .909, 48);
    stat(s, '02', 5.486, 2.804, 1.268, .909, 48);
    stat(s, '03', 5.486, 5.09, 1.268, .909, 48);
    rule(s, 5.617, 1.606, 1.826, 0);
    rule(s, 5.617, 6.199, 1.826, 0);
    rule(s, 5.617, 3.939, 1.826, 0);
    corner(s, 'C14', T.orn);
    note(s, T.sub, 10.428, 4.455, 4.315, .404, 18, { align: 'right', rotate: 270 });
    head(s, 'PRODUCT DEVELOPMENT', 8.337, 3.242, 7.06, 1.111, 60, { align: 'right', rotate: 270 });
    head(s, 'PROCESS', 11.535, 1.096, 2.101, .707, 36, { align: 'right', rotate: 270 });
  },

  /* 36 — BUSINESS GOAL ONE */
  function (s) {
    img(s, 8.59, 0, 4.744, 7.5);
    icon(s, P.pole, 1.581, .958, .207, 6.553, INK);
    icon(s, P.tagBanner, 1.923, 1.291, 4.9, .841, '737373');
    icon(s, P.tagBanner, 1.923, 4.23, 4.9, .854, 'A6A6A6');
    head(s, 'BUSINESS GOAL ONE', 2.957, 1.392, 3.549, .64, 32);
    head(s, 'BUSINESS GOAL TWO', 2.957, 4.337, 3.549, .64, 32);
    body(s, T.p1, 2.957, 2.261, 4.834, 1.583);
    body(s, T.p1, 2.957, 5.278, 4.834, 1.583);
    head(s, 'BUSINESS GOAL', 9.785, .638, 3.549, 2.121, 60);
    note(s, T.sub, 9.785, 2.777, 2.549, .707, 18);
    rule(s, 9.549, -.019, 0, 3.502);
    corner(s, 'C15', T.orn);
  },

  /* 37 — BUSINESS GOAL THREE */
  function (s) {
    img(s, 8.59, 0, 4.744, 7.5);
    icon(s, P.baseShadow, .689, 6.084, 1.956, .691, '262626');
    icon(s, P.baseTop, .689, 6.084, 1.956, .548, INK);
    icon(s, P.pole, 1.576, 0, .206, 6.657, INK);
    icon(s, P.tagBanner, 1.915, .369, 4.864, .834, 'BFBFBF');
    icon(s, P.tagBanner, 1.915, 3.402, 4.864, .848, 'D9D9D9');
    body(s, T.p1, 2.895, 1.414, 4.834, 1.583);
    body(s, T.p1, 2.895, 4.518, 4.834, 1.583);
    head(s, 'BUSINESS GOAL THREE', 2.895, .461, 3.666, .64, 32);
    head(s, 'BUSINESS GOAL FOUR', 2.895, 3.522, 3.666, .64, 32);
    corner(s, 'C14', T.orn);
    head(s, 'BUSINESS GOAL', 9.785, 4.978, 3.549, 2.121, 60);
    note(s, T.sub, 9.785, 4.131, 2.549, .707, 18);
    rule(s, 9.549, 3.998, 0, 3.502);
  },

  /* 38 — STATISTIC */
  function (s) {
    frame(s, 3.56, 3.822, 3.262, 3.262, INK);
    frame(s, 9.402, 3.822, 3.262, 3.262, INK);
    box(s, .671, 1.754, 3.262, 3.262, INK);
    head(s, 'STATISTIC', 10.444, .539, 2.162, 2.121, 60);
    note(s, T.sub, 6.835, .637, 3.148, .707, 18, { align: 'right' });
    rule(s, 10.232, 0, 0, 2.364);
    corner(s, 'C16', T.orn);
    stat(s, '100 K', .671, 2.84, 3.262, .909, 48, { color: CANVAS, align: 'center' });
    note(s, '23 K', 7.432, 3.135, 1.65, .64, 32, { fontFace: 'Arvo bold', color: CANVAS, charSpacing: 3, align: 'center' });
    stat(s, '500', 3.56, 5.103, 3.262, .909, 48, { align: 'center' });
    icon(s, P.thumb, 2.212, 2.192, .343, .469, 'EBECEE');
    box(s, 2.05, 2.364, .107, .255, 'EBECEE');
    icon(s, P.trophy, 4.916, 4.304, .57, .579, INK);
    icon(s, P.cart, 10.772, 4.353, .53, .53, INK);
    head(s, 'Satiesfied Clients', 1.115, 3.784, 2.374, .909, 24, { color: PAPER, align: 'center' });
    box(s, 6.512, 1.663, 3.262, 3.262, INK);
    stat(s, '23 K', 6.512, 2.847, 3.262, .909, 48, { color: CANVAS, align: 'center' });
    icon(s, P.targetRing, 7.866, 2.23, .478, .481, 'EBECEE');
    icon(s, P.targetArrow, 8.093, 2.158, .327, .327, 'EBECEE');
    icon(s, P.targetDot, 7.987, 2.352, .237, .237, 'EBECEE');
    head(s, 'Project Done', 6.956, 3.791, 2.374, .505, 24, { color: PAPER, align: 'center' });
    head(s, 'AWARDS MENTION', 4.056, 6.001, 2.374, .505, 24, { align: 'center' });
    stat(s, '300 K', 9.398, 5.103, 3.262, .909, 48, { align: 'center' });
    head(s, 'REPEAT ORDER', 9.846, 6.001, 2.374, .505, 24, { align: 'center' });
  },

  /* 39 — TESTIMONIALS */
  function (s) {
    head(s, 'TESTIMONIALS', 4.219, .531, 4.895, 1.111, 60, { align: 'center' });
    note(s, T.sub, 4.031, 1.712, 5.271, .404, 18, { align: 'center' });
    icon(s, P.person, 2.492, 2.535, .687, 1.943, INK);
    icon(s, P.person, 3.344, 2.535, .678, 1.943, INK);
    icon(s, P.person, 4.188, 2.535, .687, 1.943, INK);
    icon(s, P.person, 5.04, 2.535, .687, 1.943, INK);
    icon(s, P.person, 5.892, 2.535, .678, 1.943, INK);
    icon(s, P.person, 6.763, 2.535, .687, 1.943, INK);
    icon(s, P.person, 7.615, 2.535, .678, 1.943, INK);
    icon(s, P.person, 8.458, 2.535, .687, 1.943, INK);
    icon(s, P.person, 9.311, 2.535, .687, 1.943, '737373');
    icon(s, P.person, 10.163, 2.535, .678, 1.943, '737373');
    body(s, 'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. ', 3.103, 5.482, 7.127, .978);
    head(s, '8 in 10 Customers Satiesfied with Our Services', 3.396, 4.931, 6.542, .438, 20, { color: HEADLINE, align: 'center' });
    corner(s, 'C3', T.orn);
    corner(s, 'C8', '2019');
  },

  /* 40 — TESTIMONIALS */
  function (s) {
    box(s, 0, 2.807, 13.333, 1.886, CANVAS);
    body(s, [{ text: T.p11, options: { breakLine: true } }], 4.743, .777, 6.11, 1.886);
    body(s, [{ text: T.p11, options: { breakLine: true } }], 2.45, 4.869, 6.11, 1.886, { align: 'right' });
    corner(s, 'C17', '2019');
    corner(s, 'C15', T.orn);
    head(s, 'TESTIMONIALS', 4.219, 2.969, 4.895, 1.111, 60, { align: 'center' });
    note(s, T.sub, 4.031, 4.127, 5.271, .404, 18, { align: 'center' });
    img(s, 0, 0, 4.182, 3.262);
    img(s, 9.152, 4.238, 4.182, 3.262);
  },

  /* 41 — OUR TEAM */
  function (s) {
    img(s, 1.192, 0, 7.062, 7.5);
    head(s, 'OUR TEAM', 2.564, 1.497, 4.895, 1.111, 60, { align: 'center' });
    note(s, T.sub, 2.376, 2.678, 5.271, .404, 18, { align: 'center' });
    body(s, T.p1, 8.188, 4.208, 3.954, 1.886);
    corner(s, 'C3', T.orn);
  },

  /* 42 — OUR TEAM */
  function (s) {
    head(s, 'OUR TEAM', 4.219, .35, 4.895, 1.111, 60, { align: 'center' });
    note(s, T.sub, 4.031, 1.531, 5.271, .404, 18, { align: 'center' });
    head(s, 'ZAHEER KAN', 4.892, 5.753, 3.549, .707, 36, { align: 'center' });
    head(s, 'CAROLINE STUART', .879, 5.664, 3.549, .707, 36, { align: 'center' });
    head(s, 'ROSALINE KYLE', 8.906, 5.664, 3.549, .707, 36, { align: 'center' });
    note(s, 'Interior Designer', .913, 6.46, 3.481, .404, 18, { align: 'center' });
    note(s, 'Project Manager', 4.926, 6.548, 3.481, .404, 18, { align: 'center' });
    note(s, 'Community Manager', 8.94, 6.46, 3.481, .404, 18, { align: 'center' });
    img(s, .913, 2.454, 3.481, 2.986);
    img(s, 4.664, 2.229, 4.005, 3.436);
    img(s, 8.94, 2.454, 3.481, 2.986);
  },

  /* 43 — CAROLINE STUART */
  function (s) {
    head(s, 'CAROLINE STUART', 1.049, 5.713, 5.033, .841, 44);
    note(s, 'Interior Designer', -.834, 2.285, 2.954, .404, 18, { align: 'right', rotate: 270 });
    box(s, 7.252, 1.636, 4.232, .167, 'D0CECE');
    box(s, 7.256, 1.636, 3.535, .167, INK);
    box(s, 7.252, 2.757, 4.232, .167, 'D0CECE');
    box(s, 7.256, 2.757, 3.535, .167, INK);
    box(s, 7.252, 3.881, 4.232, .167, 'D0CECE');
    box(s, 7.256, 3.881, 3.535, .167, INK);
    head(s, 'Marketing Communication', 7.252, 3.306, 4.765, .404, 18, { color: HEADLINE });
    head(s, 'Public Speaking', 7.252, 2.213, 3.426, .404, 18, { color: HEADLINE });
    head(s, 'Social Branding', 7.252, 1.09, 3.426, .404, 18, { color: HEADLINE });
    rule(s, .65, 3.443, 0, 4.057);
    body(s, T.p10, 7.252, 4.572, 4.765, 1.28);
    corner(s, 'C18', T.orn);
    img(s, 1.049, .959, 5.033, 4.59);
  },

  /* 44 — ZAHEER KAN */
  function (s) {
    head(s, 'ZAHEER KAN', 1.06, 1.156, 5.011, .841, 44);
    note(s, 'Project Manager', 3.738, 1.959, 2.333, .404, 18, { align: 'right' });
    rule(s, -.016, 2.165, 3.738, .022, { flipH: true });
    body(s, T.p1, 7.262, 2.774, 4.41, 1.583);
    icon(s, P.facebook, 7.383, 5.034, .181, .365, INK);
    icon(s, P.twitter, 7.262, 5.68, .422, .35, INK);
    head(s, 'ZAHEERKAN', 8.02, 5.034, 1.841, .37, 16);
    head(s, '@ZAHEERKAN', 8.032, 5.618, 1.841, .37, 16);
    corner(s, 'C19', T.orn);
    note(s, '547 M', 9.932, 5.034, 1.82, .572, 28, { fontFace: 'Staatliches', charSpacing: 1, bold: true });
    head(s, 'Followers ', 9.92, 5.641, 1.82, .438, 20);
    img(s, 1.06, 2.689, 5.011, 4.328);
  },

  /* 45 — ROSALINE KYLE */
  function (s) {
    head(s, 'ROSALINE KYLE', 5.014, 3.161, 4.426, .841, 44, { align: 'center', rotate: 270 });
    note(s, 'Community Manager', 7.77, .742, 3.481, .404, 18);
    rule(s, 10.607, .95, 2.727, .016, { flipH: true });
    box(s, 1.487, 1.51, 4.232, .167, 'D0CECE');
    box(s, 1.491, 1.51, 3.535, .167, INK);
    box(s, 1.487, 2.713, 4.232, .167, 'D0CECE');
    box(s, 1.491, 2.713, 3.535, .167, INK);
    box(s, 1.487, 3.967, 4.232, .167, 'D0CECE');
    box(s, 1.491, 3.967, 3.535, .167, INK);
    head(s, 'Marketing Communication', 1.487, 3.196, 4.765, .404, 18, { color: HEADLINE });
    head(s, 'Public Speaking', 1.487, 2.169, 3.426, .404, 18, { color: HEADLINE });
    head(s, 'Social Branding', 1.487, .963, 3.426, .404, 18, { color: HEADLINE });
    body(s, T.p10, 1.619, 4.814, 4.232, 1.583);
    corner(s, 'C15', T.orn);
    icon(s, P.facebook, 8.116, 6.131, .181, .365, INK);
    head(s, 'ZAHEERKAN', 8.425, 6.131, 1.841, .37, 16);
    icon(s, P.twitter, 10.316, 6.169, .422, .35, INK);
    head(s, '@ZAHEERKAN', 10.906, 6.108, 1.841, .37, 16);
    img(s, 7.82, 1.369, 4.918, 4.426);
  },

  /* 46 — GALLERY SERVICES */
  function (s) {
    img(s, 1.347, .907, 4.426, 3.852);
    img(s, 7.56, 2.773, 4.426, 3.852);
    head(s, 'GALLERY SERVICES', 4.219, 3.022, 4.895, 2.121, 60, { align: 'center' });
    note(s, T.sub, 4.031, 5.142, 5.271, .404, 18, { align: 'center' });
    body(s, 'PLACEHOLDER', 1.347, 4.988, 2.684, 1.28);
    body(s, 'Aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo.', 8.799, 1.11, 3.187, 1.28, { align: 'right' });
    note(s, 'Modern', -.64, 2.599, 3.381, .404, 18, { charSpacing: 3, rotate: 270 });
    note(s, 'Minimalist', 10.756, 4.443, 3.247, .404, 18, { charSpacing: 3, rotate: 270 });
    note(s, 'Elegance', 1.623, .346, 3.852, .404, 18, { charSpacing: 3 });
    note(s, 'Trends', 7.918, 6.806, 3.77, .404, 18, { charSpacing: 3 });
  },

  /* 47 — GALLERY SERVICES */
  function (s) {
    head(s, 'GALLERY SERVICES', 7.334, 4.53, 4.895, 2.121, 60);
    note(s, T.sub, 7.334, 6.65, 4.895, .404, 18);
    body(s, T.p10, 2.651, .868, 3.447, 1.886, { align: 'right' });
    note(s, 'Minimalist', .206, 5.112, 3.481, .404, 18, { charSpacing: 3, rotate: 270 });
    note(s, 'Minimalist', 9.652, 1.988, 3.489, .404, 18, { charSpacing: 3, align: 'right', rotate: 270 });
    img(s, 2.316, 3.115, 4.585, 4.377);
    img(s, 6.433, 0, 4.585, 4.377);
  },

  /* 48 — GALLERY SERVICES */
  function (s) {
    head(s, 'GALLERY SERVICES', -2.072, 3.195, 6.492, 1.111, 60, { align: 'center', rotate: 270 });
    note(s, T.sub, 9.564, 6.254, 3.218, .707, 18, { align: 'right' });
    body(s, T.p3, 3.18, 1.049, 3.087, 1.886);
    note(s, 'Artistic', 6.538, 6.065, 1.794, .404, 18, { charSpacing: 3, align: 'center', rotate: 270 });
    note(s, 'Modern', 6.667, 1.106, 4.721, .404, 18, { charSpacing: 3, align: 'center' });
    corner(s, 'C3', T.orn);
    img(s, 2.393, 3.41, 4.721, 3.754);
    img(s, 6.667, 1.616, 4.721, 3.754);
  },

  /* 49 — THE HOME */
  function (s) {
    img(s, 0, 0, 6.667, 7.5);
    note(s, 'Furniture should always be', 6.005, .621, 5.765, 1.313, 36);
    head(s, 'THE HOME', 6.003, 4.736, 4.317, 1.212, 66);
    head(s, 'comfortable', 8.623, 1.596, 4.317, 1.01, 54);
    note(s, 'And always have a piece of art that you made somewhere in ', 5.951, 2.722, 5.765, 1.919, 36);
    note(s, 'Tamara Taylor', 10.126, 6.112, 2.596, .505, 24, { align: 'right' });
    rule(s, 10.311, 6.848, 3.022, 0);
  },

  /* 50 — THANKS */
  function (s) {
    img(s, 0, 0, 13.333, 7.5);
    head(s, 'THANKS', 6.667, 2.977, 5.874, 1.717, 96, { align: 'center' });
    note(s, 'Your Kind Word Here', 7.347, 4.824, 4.514, .404, 18, { align: 'center' });
  },

];

/* -------------------------------------------------------------------- build */

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';         // 13.333in x 7.5in
pptx.title = 'Ornament — Presentation Template';

SLIDES.forEach(function (buildSlide, i) {
  const s = pptx.addSlide();
  s.background = { color: WHITE_SLIDES.indexOf(i + 1) === -1 ? CANVAS : PAPER };
  buildSlide(s);
});

pptx.writeFile({ fileName: path.join(__dirname, '0e07a616-0c2f-4605-9b29-085cbac6e66a_grok_final.pptx') })
  .then(function (f) { console.log('wrote ' + f); });
