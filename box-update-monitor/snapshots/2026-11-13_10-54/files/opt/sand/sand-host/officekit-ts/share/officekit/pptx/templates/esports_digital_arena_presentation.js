/**
 * Digital Arena — Esport Presentation Template (20 slides, 13.333in x 7.5in)
 * Standalone pptxgenjs recreation of the reference deck.
 *
 * Raster photos in the original are replaced by flat grey placeholder rectangles
 * (`imagePlaceholder`), per the conversion brief.
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ── Theme ──────────────────────────────────────────────────────────────────
const RED = 'FA4655';        // accent1
const INK = '292929';        // accent2
const BLACK = '000000';      // tx1
const TEXT = '262626';       // tx1 lumMod 85% / lumOff 15% — body copy grey
const WHITE = 'FFFFFF';      // bg1
const RED_LT = 'FC9099';     // accent1 lumMod 60% / lumOff 40%
const RED_DK = '9C0411';     // accent1 lumMod 50%
const RED_MD = 'EA0619';     // accent1 lumMod 75%
const INK_DK = '141414';     // accent2 lumMod 50%
const INK_LT = '7F7F7F';     // accent2 lumMod 60% / lumOff 40%
const POLE = 'DBDBDB';       // grey flag pole on slide 18
const PHOTO = 'CECECE';      // average colour of the replaced photograph
const PHOTO_TXT = 'B4B4B4';

const HEAD = 'Anton';        // major latin
const BODY = 'Work Sans';    // minor latin

const NONE = { type: 'none' };
// pptxgenjs mutates the shadow object it is handed, so hand out a fresh one per shape.
const cardShadow = () => ({ type: 'outer', color: BLACK, opacity: 0.2, blur: 50, offset: 20, angle: 45 });

// Lorem strings reused throughout the deck
const L_LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et';
const L_MED = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut';
const L_SHORT = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod';
const L_TEMPOR = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor';
const L_ADIPI = 'Lorem ipsum dolor sit amet, consectetuer adipi scing elit.';
const L_ADIPI2 = 'Lorem ipsum dolor sit amet, consectetuer adipi scing elit. Lorem ipsum dolor sit ';
const L_MAECENAS = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus posuere, magna sed pulvinar ultricies, purus lectus';
const L_STEP = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue porttitor massa on. ';
const L_CEASK = 'Lorem ipsum dolor sit amet, consectetuer adipiscing ceask';

const ICONS = {
  gamepad: [[[73,100,1],[55,87,66,100,60,95],[44,87],[26,100,40,95,33,100],[0,51,11,100,0,77],[26,0,0,23,11,0],[73,0],[100,51,88,0,100,23],[73,100,100,77,88,100],[],[42,44,1],[41,41,42,41,42,41],[32,41],[32,23],[30,18],[23,18],[22,23,22,18,22,21],[22,41],[11,41],[10,44,11,41,10,41],[10,56],[11,59,10,59,11,59],[22,59],[22,77],[23,82,22,79,22,82],[30,82],[32,77],[32,59],[41,59],[42,56,42,59,42,59],[42,44],[],[66,51,1],[59,62,63,51,59,56],[66,74,59,69,63,74],[73,62,70,74,73,69],[66,51,73,56,70,51],[],[79,26,1],[73,38,75,26,73,31],[79,51,73,44,75,51],[86,38,84,51,86,44],[79,26,86,31,84,26],[]]],
  sun: [[[48,26,1],[26,48,35,26,26,35],[48,74,26,65,35,74],[74,48,61,74,74,65],[48,26,74,35,61,26],[],[48,65,1],[32,48,42,65,32,58],[48,35,32,42,42,35],[65,48,58,35,65,42],[48,65,65,58,58,65],[],[55,10,1],[48,16,55,13,52,16],[45,10,45,16,45,13],[45,3],[48,0,45,0,45,0],[55,3,52,0,55,0],[55,10],[],[26,26,1],[19,26],[13,19],[13,13],[19,13],[26,19],[26,26],[],[13,48,1],[10,55,13,52,13,55],[3,55],[0,48,0,55,0,52],[3,45,0,48,0,45],[10,45],[13,48,13,45,13,48],[],[26,74,1],[26,81],[19,84],[13,84,19,87,16,87],[13,81],[19,74],[26,74],[],[55,97,1],[48,100,55,97,52,100],[45,97,45,100,45,97],[45,87],[48,84,45,87,45,84],[55,87,52,84,55,87],[55,97],[],[74,74,1],[81,74],[84,81],[84,84,87,81,87,84],[77,84,84,87,81,87],[74,81],[74,74,71,77,71,77],[],[100,48,1],[94,55,100,52,97,55],[87,55],[84,48,87,55,84,52],[87,45],[94,45],[100,48,97,45,100,48],[],[74,26,1],[74,19,71,23,71,19],[77,13],[84,13],[84,19,87,16,87,19],[81,26],[74,26],[]]],
  monitor: [[[93,0,1],[67,0],[52,0],[7,0],[0,9,4,0,0,5],[0,59],[7,68,0,64,4,68],[41,68],[41,77],[26,77],[19,86,22,77,19,82],[19,91],[26,100,19,95,22,100],[52,100],[67,100],[74,100],[81,91,78,100,81,95],[81,86],[74,77,81,82,78,77],[67,77],[59,77],[59,68],[67,68],[93,68],[100,59,96,68,100,64],[100,9],[93,0,100,5,96,0],[],[93,55,1],[89,64,93,59,93,64],[67,64],[52,64],[15,64],[7,55,11,64,7,59],[7,14],[15,5,7,9,11,5],[52,5],[67,5],[89,5],[93,14,93,5,93,9],[93,55],[]]],
  chat: [[[56,5,1],[16,5],[0,27,6,5,0,14],[0,55],[16,73,0,68,6,73],[9,100,16,82,9,100],[34,73,9,100,28,82],[56,73],[69,55,63,73,69,68],[69,27],[56,5,69,14,63,5],[]],[[88,0,1],[53,0],[50,0],[60,0],[75,23,69,0,75,9],[75,55],[75,59],[91,78,81,64,91,78],[84,55,91,78,88,59],[88,55],[100,36,94,55,100,46],[100,14],[88,0,100,5,94,0],[]]],
  envelope: [[[95,0,1],[94,0],[6,0],[50,64],[95,0],[]],[[100,7,1],[68,51],[99,96],[100,90],[100,10],[100,7],[]],[[0,5,1],[0,9],[0,91],[1,95],[32,51],[0,5],[]],[[50,78,1],[37,59],[6,100],[94,100],[63,59],[50,78],[]]],
  barchart: [[[6,79,1],[25,79],[25,100],[6,100],[]],[[32,65,1],[50,65],[50,100],[32,100],[]],[[56,51,1],[75,51],[75,100],[56,100],[]],[[81,37,1],[100,37],[100,100],[81,100],[]],[[87,15,1],[67,15],[43,35],[32,28],[0,56],[0,65],[32,36],[44,43],[70,22],[89,22],[100,9],[100,0],[87,15],[]]],
  microscope: [[[85,94,1],[80,94],[100,63,92,87,100,76],[68,27,100,45,86,30],[75,13],[75,9,76,11,75,9],[55,1],[52,0],[50,2],[27,40],[30,49,25,43,27,47],[27,54],[40,59],[43,54],[53,51,47,56,51,54],[61,38],[85,63,75,40,85,50],[57,88,85,77,73,88],[36,82,50,88,41,85],[36,78],[39,75,36,77,37,75],[57,75],[57,69],[29,69],[16,69],[0,69],[0,75],[17,75],[18,75],[22,78,20,75,22,77],[22,82],[22,94],[8,100,14,94,8,93],[100,100],[85,94,100,93,93,94],[],[61,9,1],[59,10],[42,38],[36,35],[36,34],[52,9],[54,7],[57,7,55,6,56,6],[61,9],[]]],
  sitemap: [[[81,63,1],[81,44],[56,44],[56,38],[75,38],[75,0],[31,0],[31,38],[50,38],[50,44],[19,44],[19,63],[0,63],[0,100],[44,100],[44,63],[25,63],[25,50],[75,50],[75,63],[56,63],[56,100],[100,100],[100,63],[81,63],[],[38,69,1],[38,75],[6,75],[6,69],[38,69],[],[38,12,1],[38,6],[69,6],[69,12],[38,12],[],[94,75,1],[63,75],[63,69],[94,69],[94,75],[]]],
  trophy: [[[100,32,1],[69,58,100,44,87,58],[63,65],[58,76,60,68,58,71],[66,84,58,79,60,84],[78,94,72,84,78,87],[78,98],[76,100,78,100,76,100],[25,100],[24,98,24,100,24,100],[24,94],[36,84,24,87,30,84],[43,76,40,84,43,79],[39,65,43,71,42,68],[33,58,37,65,36,61],[0,32,15,58,0,44],[0,24],[6,18,0,19,3,18],[24,18],[24,11],[33,0,24,5,28,0],[67,0],[78,11,73,0,78,5],[78,18],[94,18],[100,24,99,18,100,19],[100,32],[],[24,26,1],[9,26],[9,32],[28,50,9,39,16,47],[24,26,25,44,24,35],[],[93,26,1],[78,26],[73,50,78,35,76,44],[93,32,85,47,93,39],[93,26],[]]],
  person: [[[97,77,1],[75,65,95,72,84,69],[63,59],[63,50],[67,39,63,50,66,47],[72,33,70,39,72,35],[69,26,72,31,72,25],[70,16,70,22,70,17],[50,0,70,8,62,0],[30,16,39,0,31,8],[31,26],[29,33,29,25,28,31],[33,39,29,35,30,39],[38,50,34,47,38,50],[38,59],[26,65],[3,77,16,69,6,72],[1,100,0,80,1,100],[99,100],[97,77,99,100,100,80],[]]],
  building: [[[50,25,1],[50,0],[6,0],[6,93],[0,93],[0,100],[100,100],[100,25],[50,25],[],[38,93,1],[19,93],[19,81],[38,81],[38,93],[],[44,68,1],[13,68],[13,62],[44,62],[44,68],[],[44,56,1],[13,56],[13,50],[44,50],[44,56],[],[44,43,1],[13,43],[13,38],[44,38],[44,43],[],[44,31,1],[13,31],[13,25],[44,25],[44,31],[],[44,19,1],[13,19],[13,13],[44,13],[44,19],[],[75,87,1],[63,87],[63,75],[75,75],[75,87],[],[75,68,1],[63,68],[63,56],[75,56],[75,68],[],[75,50,1],[63,50],[63,38],[75,38],[75,50],[],[94,87,1],[81,87],[81,75],[94,75],[94,87],[],[94,68,1],[81,68],[81,56],[94,56],[94,68],[],[94,50,1],[81,50],[81,38],[94,38],[94,50],[]]],
  bookClosed: [[[98,17,1],[90,17],[76,17],[14,17],[6,11,10,17,6,14],[14,5,6,8,10,5],[98,5],[100,2,100,5,100,3],[98,0,100,0,100,0],[14,0],[0,11,6,0,0,5],[0,12],[0,89],[14,100,0,95,6,100],[98,100],[100,98,100,100,100,100],[100,20],[98,17,100,19,100,17],[],[16,95,1],[14,95],[6,89,10,95,6,92],[6,20],[14,22,8,22,12,22],[16,22],[16,95],[],[94,95,1],[24,95],[24,22],[33,22],[33,23],[33,77],[35,80,33,78,33,80],[39,77,37,80,39,78],[39,23],[39,22],[49,22],[49,61],[53,64,49,62,51,64],[55,61,55,64,55,62],[55,22],[76,22],[90,22],[94,22],[94,95],[]],[[98,8,1],[14,8],[12,11,12,8,12,9],[14,14,12,12,12,14],[98,14],[100,11,100,14,100,12],[98,8,100,9,100,8],[]]],
  bookOpen: [[[99,13,1],[97,11,99,11,99,11],[92,11],[92,2],[90,0,92,0,91,0],[87,0],[49,8,73,0,56,2],[13,0,44,2,27,0],[10,0],[8,2,9,0,8,0],[8,11],[3,11],[1,13,1,11,1,11],[0,15,0,13,0,13],[0,87],[3,90,0,89,1,90],[5,90],[45,98,31,90,45,97],[47,100],[53,100],[55,98,53,100,53,98],[95,90,55,97,69,90],[97,90],[100,87,99,90,100,89],[100,15],[99,13,100,13,100,13],[],[4,85,1],[4,16],[8,16],[8,85],[5,85],[4,85],[],[12,85,1],[12,15],[12,5],[13,5],[48,13,27,5,43,7],[48,93],[12,85,45,90,30,85],[],[52,93,1],[52,13],[87,5,57,7,73,5],[88,5],[88,15],[88,85],[52,93,70,85,55,90],[],[96,85,1],[95,85],[92,85],[92,16],[96,16],[96,85],[]]],
  clipboard: [[[88,10,1],[82,10],[78,11,80,10,78,11],[82,14,78,13,80,14],[88,14],[94,19,92,14,94,16],[94,90],[88,97,94,94,92,97],[12,97],[6,90,8,97,6,94],[6,19],[12,14,6,16,8,14],[20,14],[22,11,20,14,22,13],[20,10],[12,10],[0,19,6,10,0,14],[0,90],[12,100,0,97,6,100],[88,100],[100,90,94,100,100,97],[100,19],[88,10,100,14,94,10],[]],[[34,24,1],[68,24],[76,18,72,24,76,21],[76,6],[68,0,76,3,72,0],[34,0],[26,6,28,0,26,3],[26,18],[34,24,26,21,28,24],[],[30,6,1],[34,5,30,5,32,5],[68,5],[70,6,68,5,70,5],[70,18],[68,19,70,19,68,19],[34,19],[30,18,32,19,30,19],[30,6],[]],[[80,37,1],[20,37],[18,39,20,37,18,37],[20,40,18,40,20,40],[80,40],[84,39,82,40,84,40],[80,37,84,37,82,37],[]],[[80,48,1],[20,48],[18,50,20,48,18,48],[20,51,18,51,20,51],[80,51],[84,50,82,51,84,51],[80,48,84,48,82,48],[]],[[80,60,1],[20,60],[18,61],[20,63,18,63,20,63],[80,63],[84,61,82,63,84,63],[80,60,84,61,82,60],[]],[[80,71,1],[20,71],[18,72],[20,74,18,74,20,74],[80,74],[84,72,82,74,84,74],[80,71,84,72,82,71],[]],[[80,82,1],[20,82],[18,84],[20,85,18,85,20,85],[80,85],[84,84,82,85,84,85],[80,82,84,84,82,82],[]]],
  house: [[[51,19,1],[14,56],[14,100],[40,100],[40,81],[51,69,40,75,43,69],[60,81,57,69,60,75],[60,100],[86,100],[86,56],[51,19]],[[80,31,1],[80,6],[69,6],[69,19],[51,0],[0,53],[6,59],[51,13],[94,59],[100,53],[80,31],[]]],
  dollar: [[[48,0,1],[0,52,22,0,0,22],[48,100,0,78,22,100],[100,52,78,100,100,78],[48,0,100,22,78,0],[],[48,96,1],[0,52,22,96,0,78],[48,4,0,22,22,4],[96,52,74,4,96,22],[48,96,96,78,74,96],[]],[[41,37,1],[44,41,41,41,44,41],[44,29],[41,37,44,29,41,33],[]],[[52,56,1],[52,67],[59,63,56,67,59,67],[52,56,59,59,56,56],[]],[[48,7,1],[4,52,26,7,4,26],[48,93,4,74,26,93],[93,52,74,93,93,74],[48,7,93,26,74,7],[],[52,78,1],[52,85],[45,85],[45,78],[26,59,33,78,26,70],[41,59],[45,67,41,63,41,67],[45,56],[26,37,33,52,26,48],[45,22,26,26,37,22],[45,15],[52,15],[52,22],[71,37,63,22,71,26],[59,37],[52,30,59,33,56,30],[52,41],[74,59,63,44,74,48],[52,78,74,70,63,78],[]]],
  globe: [[[100,48,1],[53,0,100,21,80,0],[50,0],[47,0],[3,48,23,0,3,21],[0,48],[0,52],[0,55],[3,55],[47,100,3,79,23,100],[50,100],[53,100],[100,55,80,100,100,79],[100,52],[100,48],[],[33,10,1],[23,28,30,14,27,21],[13,28],[33,10,20,21,27,14],[],[10,34,1],[23,34],[20,48,23,38,20,41],[7,48],[10,34,10,41,10,38],[],[7,55,1],[20,55],[23,69,20,59,23,62],[10,69],[7,55,10,62,10,59],[],[13,76,1],[23,76],[33,90],[13,76,27,86,20,83],[],[47,93,1],[30,76,40,93,37,86],[47,76],[47,93],[],[47,69,1],[30,69],[27,55,30,62,27,59],[47,55],[47,69],[],[47,48,1],[27,48],[30,34,27,41,30,38],[47,34],[47,48],[],[47,28,1],[30,28],[47,7,37,17,40,10],[47,28],[],[87,28,1],[77,28],[70,10,77,21,73,14],[87,28,77,14,83,21],[],[53,7,1],[70,28,60,10,67,17],[53,28],[53,7],[],[53,34,1],[73,34],[73,48],[53,48],[53,34],[],[53,55,1],[73,55],[73,69],[53,69],[53,55],[],[53,93,1],[53,76],[70,76],[53,93,67,86,60,93],[],[70,90,1],[77,76,73,86,77,79],[87,76],[70,90,83,83,77,86],[],[90,69,1],[80,69],[80,55],[93,55],[90,69,93,59,93,62],[],[80,48,1],[80,34],[90,34],[93,48,93,38,93,41],[80,48],[]]],
  yen: [[[48,0,1],[0,52,22,0,0,22],[48,100,0,78,22,100],[100,52,78,100,100,78],[48,0,100,22,78,0],[],[48,96,1],[4,52,22,96,4,78],[48,4,4,26,22,4],[96,52,74,4,96,26],[48,96,96,78,74,96],[]],[[48,7,1],[8,52,26,7,8,26],[48,92,8,74,26,92],[93,52,74,92,93,74],[48,7,93,26,74,7],[],[63,55,1],[63,63],[56,63],[56,66],[63,66],[63,70],[56,70],[56,81],[45,81],[45,70],[34,70],[34,66],[45,66],[45,63],[34,63],[34,55],[41,55],[26,26],[41,26],[45,41],[48,52,48,44,48,48],[56,41,52,48,52,44],[59,26],[74,26],[56,55],[63,55],[]]],
  euro: [[[50,0,1],[0,52,21,0,0,22],[50,100,0,78,21,100],[100,52,75,100,100,78],[50,0,100,22,75,0],[],[50,96,1],[4,52,25,96,4,78],[50,4,4,26,25,4],[96,52,75,4,96,26],[50,96,96,78,75,96],[]],[[50,7,1],[7,52,25,7,7,26],[50,92,7,74,25,92],[93,52,71,92,93,74],[50,7,93,26,71,7],[],[64,33,1],[60,33],[57,33],[39,44,50,33,43,37],[46,44],[46,52],[39,52],[46,52],[46,59],[39,59],[57,70,43,66,50,70],[60,70],[64,66],[64,77],[60,81],[57,81],[28,59,43,81,32,70],[25,59],[25,52],[28,52],[25,52],[25,44],[28,44],[57,22,32,29,43,22],[60,22],[64,22],[64,33],[]]],
  gem: [[[100,25,1],[86,2],[85,0,86,0,85,0],[51,0],[49,0],[15,0],[12,2],[0,25],[0,26],[0,28],[46,98],[48,98],[48,100],[49,100],[51,100],[51,98],[100,28],[100,26],[100,25],[],[49,9,1],[58,23],[42,23],[49,9],[],[17,5,1],[46,5],[35,23],[8,23],[17,5],[],[8,29,1],[34,29],[43,83],[8,29],[],[49,82,1],[40,29],[60,29],[49,82],[],[55,83,1],[66,29],[92,29],[55,83],[],[65,23,1],[54,5],[83,5],[92,23],[65,23],[]]],
  docChart: [[[67,0,1],[0,0],[0,100],[100,100],[100,28],[67,0],[],[85,87,1],[15,87],[15,12],[60,12],[85,34],[85,87],[]],[[22,69,1],[35,69],[35,81],[22,81],[]],[[43,37,1],[57,37],[57,81],[43,81],[]],[[64,56,1],[78,56],[78,81],[64,81],[]]],
  palette: [[[56,0,1],[0,66,19,0,0,39],[19,66,0,73,13,86],[38,73,22,58,38,60],[56,100,38,88,44,100],[100,47,94,100,100,66],[56,0,100,27,94,0],[],[25,48,1],[17,39,21,48,17,44],[25,31,17,35,21,31],[33,39,29,31,33,35],[25,48,33,44,29,48],[],[39,23,1],[47,15,39,18,43,15],[55,23,52,15,55,18],[47,31,55,28,52,31],[39,23,43,31,39,28],[],[59,82,1],[48,70,53,82,48,76],[59,58,48,63,53,58],[71,70,66,58,71,63],[59,82,71,76,66,82],[],[78,61,1],[73,57,76,61,73,60],[78,51,73,54,76,51],[83,57,81,51,83,54],[78,61,83,60,81,61],[],[72,41,1],[61,30,66,41,61,37],[72,18,61,24,66,18],[83,30,78,18,83,24],[72,41,83,37,78,41],[]]],
  gift: [[[6,62,1],[47,62],[47,100],[6,100],[]],[[53,62,1],[94,62],[94,100],[53,100],[]],[[60,32,1],[60,31],[86,16,70,30,86,27],[70,0,86,8,78,0],[53,25,62,0,55,7],[47,25],[30,0,45,7,38,0],[14,16,22,0,14,8],[40,31,14,27,30,30],[40,32],[0,32],[0,56],[47,56],[47,32],[53,32],[53,56],[100,56],[100,32],[60,32],[],[20,16,1],[30,7,20,11,26,7],[39,25,35,7,39,14],[20,16,30,24,20,21],[],[70,7,1],[80,16,74,7,80,11],[61,25,80,21,70,24],[70,7,61,14,65,7],[]]],
  phone: [[[93,14,1],[50,0,85,6,75,0],[7,14,25,0,15,6],[0,26,3,17,0,20],[0,33],[6,39,0,36,3,39],[19,39],[25,33,22,39,25,36],[28,22,25,29,26,26],[50,14,32,18,38,13],[72,22,63,13,69,18],[75,33,75,26,75,29],[82,39,75,36,78,39],[94,39],[100,33,97,39,100,36],[100,26],[93,14,100,20,97,17],[]],[[38,60,1],[63,60],[63,86],[38,86],[]],[[71,40,1],[69,40],[69,31],[63,27,69,29,66,27],[56,31,59,27,56,29],[56,40],[44,40],[44,31],[38,27,44,29,41,27],[31,31,34,27,31,29],[31,40],[28,40],[23,43,26,40,24,41],[6,80,18,51,6,71],[6,93],[13,100,6,96,9,100],[88,100],[94,93,91,100,94,96],[94,81],[76,42,94,69,82,50],[71,40,76,41,74,40],[],[50,93,1],[31,73,39,93,31,84],[50,53,31,62,39,53],[69,73,60,53,69,62],[50,93,69,84,60,93],[]]],
  qr: [[[47,45,1],[0,45],[0,0],[47,0],[47,45],[],[47,100,1],[0,100],[0,55],[47,55],[47,100],[],[37,37,1],[37,8],[9,8],[9,37],[37,37],[],[37,90,1],[37,64],[9,64],[9,90],[37,90],[],[28,27,1],[18,27],[18,17],[28,17],[28,27],[],[28,82,1],[18,82],[18,72],[28,72],[28,82],[],[100,45,1],[56,45],[56,0],[100,0],[100,45],[],[100,82,1],[72,82],[72,72],[63,72],[63,100],[56,100],[56,55],[82,55],[82,64],[91,64],[91,55],[100,55],[100,82],[],[91,37,1],[91,8],[63,8],[63,37],[91,37],[],[82,27,1],[72,27],[72,17],[82,17],[82,27],[],[82,100,1],[72,100],[72,90],[82,90],[82,100],[],[100,100,1],[91,100],[91,90],[100,90],[100,100],[]]],
  notebook: [[[98,33,1],[83,90],[70,100,81,95,77,100],[16,100],[2,88,9,100,3,95],[2,80,0,85,0,82],[2,77],[2,73],[3,68],[8,57,6,67,8,60],[8,53],[11,50],[14,38,12,48,14,42],[14,35],[17,32,14,33,16,33],[20,20,19,30,20,23],[20,17],[23,13],[34,3,25,10,27,0],[36,3],[81,3],[89,7,84,3,88,3],[89,15,91,8,91,12],[73,72],[61,83,70,82,69,83],[9,83],[8,85,9,83,8,83],[8,87],[16,92,9,92,12,92],[70,92],[77,88,73,92,75,90],[94,25],[94,22],[98,25],[98,33,100,27,100,30],[],[31,43,1],[67,43],[70,42,69,43,69,42],[70,37],[70,35,72,37,70,35],[34,35],[31,37,33,35,31,37],[30,42],[31,43,30,42,30,43],[],[36,27,1],[72,27],[75,25,73,27,75,27],[77,20],[75,18],[39,18],[36,20,38,18,36,20],[34,25],[36,27,34,27,34,27],[]]],
};

// ── Helpers ────────────────────────────────────────────────────────────────

/** Solid rectangle (or any preset shape). */
function rect(s, x, y, w, h, opts) {
  s.addShape('rect', Object.assign({ x, y, w, h, line: NONE }, opts));
}

/** Preset shape by name. */
function shape(s, name, x, y, w, h, opts) {
  s.addShape(name, Object.assign({ x, y, w, h, line: NONE }, opts));
}

/**
 * Text box. `runs` is either a string or an array of { text, options } run objects.
 * Layout defaults match the deck: Work Sans, top-anchored, no bullet, zero inset.
 */
function text(s, runs, x, y, w, h, opts) {
  s.addText(runs, Object.assign({
    x, y, w, h,
    fontFace: BODY, fontSize: 10.5, color: TEXT,
    align: 'left', valign: 'top', bullet: false, wrap: true, isTextBox: true,
  }, opts));
}

/** The heavy Anton display type used for every slide headline. */
function heading(s, runs, x, y, w, h, opts) {
  text(s, runs, x, y, w, h, Object.assign({ fontFace: HEAD, fontSize: 40 }, opts));
}

/** Body copy: 10.5pt Work Sans, justified, 150% leading. */
function para(s, str, x, y, w, h, opts) {
  text(s, str, x, y, w, h, Object.assign({ align: 'justify', lineSpacingMultiple: 1.5 }, opts));
}

/** Grey block standing in for a photograph in the source deck. */
function imagePlaceholder(s, x, y, w, h, opts) {
  rect(s, x, y, w, h, Object.assign({ fill: { color: PHOTO } }, opts));
  if (!opts || !opts.noCaption) {
    text(s, '[image]', x, y + h / 2 - 0.18, w, 0.36, { align: 'center', color: PHOTO_TXT, fontSize: 12 });
  }
}

/**
 * Draw a vector icon from ICONS. Each icon is a list of sub-paths; each sub-path is a
 * list of segments in a 0..100 box: [x,y,1]=moveTo, [x,y]=lineTo,
 * [x,y,cx1,cy1,cx2,cy2]=cubic, []=close.
 */
function icon(s, key, x, y, w, h, color) {
  ICONS[key].forEach(sub => {
    // Point coordinates are relative to the shape frame.
    const pts = sub.map(p => {
      if (p.length === 0) return { close: true };
      const px = (p[0] / 100) * w, py = (p[1] / 100) * h;
      if (p.length === 6) {
        return { x: px, y: py, curve: { type: 'cubic',
          x1: (p[2] / 100) * w, y1: (p[3] / 100) * h,
          x2: (p[4] / 100) * w, y2: (p[5] / 100) * h } };
      }
      return { x: px, y: py, moveTo: p.length === 3 };
    });
    s.addShape('custGeom', { x, y, w, h, points: pts, fill: { color }, line: NONE });
  });
}

/** Freeform polygon/curve given directly in 0..100 space (same encoding as icons). */
function poly(s, segs, x, y, w, h, opts) {
  const pts = segs.map(p => {
    if (p.length === 0) return { close: true };
    const px = (p[0] / 100) * w, py = (p[1] / 100) * h;
    if (p.length === 6) {
      return { x: px, y: py, curve: { type: 'cubic',
        x1: (p[2] / 100) * w, y1: (p[3] / 100) * h,
        x2: (p[4] / 100) * w, y2: (p[5] / 100) * h } };
    }
    return { x: px, y: py, moveTo: p.length === 3 };
  });
  s.addShape('custGeom', Object.assign({ x, y, w, h, points: pts, line: NONE }, opts));
}

/** Top navigation bar (gamepad mark + menu words + "Game" wordmark) — on every slide. */
function navBar(s, tone, wordmark) {
  const c = tone || BLACK;
  icon(s, 'gamepad', 1.159, 0.413, 0.444, 0.236, c);
  [['About us', 3.761, 0.851], ['Albums', 5.217, 0.745], ['Gallery', 6.567, 0.703]].forEach(([label, x, w]) => {
    text(s, label, x, 0.392, w, 0.278, { color: c, wrap: false });
  });
  text(s, 'Game', 11.533, 0.329, 0.761, 0.404,
    { fontFace: HEAD, fontSize: 18, color: (wordmark || c), wrap: false });
}

/** Two-tone Anton headline: dark first line/word, red second. */
function twoToneHeading(s, first, second, x, y, w, h, opts) {
  heading(s, [
    { text: first, options: { color: (opts && opts.firstColor) || INK, breakLine: !(opts && opts.inline) } },
    { text: second, options: { color: RED } },
  ], x, y, w, h, opts);
}

/** The "Infographic Section" title shared by slides 9-18. */
function sectionTitle(s) {
  heading(s, [
    { text: 'Infographic ', options: { color: INK } },
    { text: 'Section', options: { color: RED } },
  ], 3.111, 1.318, 7.111, 0.774, { align: 'center' });
}

// ── Slides ─────────────────────────────────────────────────────────────────

// 1 — Title: "Digital Arena"
function slide01(p) {
  const s = p.addSlide();
  // Gradient photo-grey → ink down the lower band, built from thin bands
  // because pptxgenjs cannot emit a real gradFill.
  const BANDS = 64;
  for (let i = 0; i < BANDS; i++) {
    const t = i / (BANDS - 1);
    const v = Math.round(0xC6 + (0x2B - 0xC6) * t);
    const hex = v.toString(16).toUpperCase().padStart(2, '0').repeat(3);
    rect(s, 0, 5.304 + (2.196 / BANDS) * i, 13.333, 2.196 / BANDS + 0.01, { fill: { color: hex } });
  }
  imagePlaceholder(s, 7.899, 1.39, 5.072, 4.774);

  heading(s, [
    { text: 'Digital ', options: { color: INK } },
    { text: 'Arena', options: { color: RED } },
  ], 1.018, 1.336, 5.878, 1.447, { fontSize: 80, valign: 'middle' });
  text(s, 'Esport Presentation Template', 1.018, 2.675, 3.918, 0.404, { fontSize: 18, wrap: false });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam tincidunt ante nec sem congue ' +
    'convallis. Pellentesque vel mauris quis nisl ornare rutrum in id risus. Proin vehicula ut sem et tempus. ' +
    'Interdum et malesuada fames ac ante ipsum primis in faucibus. ',
    0.869, 3.55, 6.083, 1.167, { lineSpacingMultiple: 1.3, valign: 'middle', margin: [19.2, 19.2, 9.6, 9.6] });

  rect(s, 1.159, 5.762, 2.029, 0.348, { fill: { color: RED } });
  text(s, 'Registration', 1.159, 5.762, 2.029, 0.348, { fontSize: 12, color: WHITE, align: 'center', valign: 'middle' });
  rect(s, 3.188, 5.762, 3.479, 0.348, { fill: { color: INK } });
  s.addShape('line', { x: 3.528, y: 5.944, w: 2.681, h: 0,
    line: { color: WHITE, width: 1, endArrowType: 'triangle' } });

  navBar(s);
  // Scattered rotated squares over the photo
  rect(s, 7.779, 1.461, 0.751, 0.751, { fill: { color: RED }, rotate: -25.4 });
  rect(s, 7.745, 4.181, 0.583, 0.583, { fill: NONE, line: { color: RED, width: 1 }, rotate: 11.1 });
  rect(s, 10.88, 5.162, 0.306, 0.306, { fill: { color: RED }, rotate: -9 });
  rect(s, 12.813, 3.249, 0.306, 0.306, { fill: { color: RED }, rotate: -38.2 });
  rect(s, 12.197, 1.698, 0.583, 0.583, { fill: NONE, line: { color: RED, width: 1 }, rotate: -46.3 });
  rect(s, 12.781, 1.441, 0.245, 0.245, { fill: NONE, line: { color: RED, width: 1 }, rotate: -46.3 });
}

// 2 — "Welcome To Digital Arena"
function slide02(p) {
  const s = p.addSlide();
  rect(s, 1.162, 1.39, 3.29, 4.774, { fill: NONE, line: { color: RED, width: 1 } });
  rect(s, 4.457, 1.057, 1.784, 1.684, { fill: NONE, line: { color: RED, width: 1 } });
  rect(s, 4.624, 1.872, 1.784, 1.684, { fill: { color: WHITE } });
  imagePlaceholder(s, 1.972, 2.167, 4.269, 4.018);

  twoToneHeading(s, 'Welcome To', 'Digital Arena', 7.344, 1.294, 3.259, 1.447);
  navBar(s);
  para(s, L_LONG, 7.346, 3.046, 4.931, 0.603);
  text(s, 'About us', 7.357, 4.288, 1.189, 0.37, { fontSize: 16, bold: true, wrap: false });
  para(s, L_LONG, 7.35, 4.81, 4.906, 0.601);

  rect(s, 9.698, 5.903, 2.494, 0.209, { fill: { color: WHITE }, line: { color: RED, width: 1 } });
  rect(s, 7.458, 5.903, 3.625, 0.209, { fill: { color: RED }, line: { color: RED, width: 1 } });
  rect(s, 1.866, 5.501, 1.012, 1.012, { fill: { color: RED }, rotate: -30 });
  rect(s, 5.521, 1.983, 0.609, 0.609, { fill: { color: RED }, rotate: 30 });
}

// 3 — "Vision & Mission"
function slide03(p) {
  const s = p.addSlide();
  navBar(s);
  rect(s, 1.159, 1.362, 11.015, 2.386, { fill: { color: WHITE }, shadow: cardShadow() });
  imagePlaceholder(s, 1.159, 1.343, 5.508, 2.407);

  heading(s, [
    { text: 'Vision & ', options: { color: INK } },
    { text: 'Mission', options: { color: RED } },
  ], 7.159, 1.824, 4.238, 0.774, { align: 'center' });
  para(s, L_TEMPOR, 7.34, 2.707, 4.132, 0.601);

  s.addShape('line', { x: 1.159, y: 4.615, w: 11.015, h: 0,
    line: { color: RED, width: 1, endArrowType: 'triangle' } });

  [['Vision', 1.057, 1.043], ['Mission', 7.354, 7.34]].forEach(([label, lx, bx]) => {
    text(s, label, lx, 4.288, 1.1, 0.37, { fontSize: 16, bold: true, wrap: false });
    [4.839, 5.621].forEach(by => {
      para(s, L_MED, bx, by, 4.95, 0.601, { bullet: { indent: 13.5 } });
    });
  });
}

// 4 — "Digital Arena Facilities" — three cards
const S4_CARDS = [
  { x: 1.159, y: 4.516, label: '01. Facilities', dark: true },
  { x: 5.054, y: 4.49, label: '02. Facilities', dark: false },
  { x: 8.965, y: 4.49, label: '03. Facilities', dark: true },
];
function slide04(p) {
  const s = p.addSlide();
  navBar(s);
  heading(s, [
    { text: 'Digital Arena', options: { color: BLACK, breakLine: true } },
    { text: 'Facilities', options: { color: RED } },
  ], 1.035, 1.306, 3.138, 1.447);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor Lorem ipsum dolor sit',
    1.035, 2.956, 3.374, 0.867);
  imagePlaceholder(s, 6.677, 1.396, 6.656, 2.354);

  S4_CARDS.forEach(c => {
    const fg = c.dark ? WHITE : BLACK;
    rect(s, c.x, c.y, 3.225, 1.614, { fill: { color: c.dark ? RED : WHITE },
      shadow: c.dark ? undefined : cardShadow() });
    text(s, c.label, c.x + 0.423, c.y + 0.174, 2.379, 0.37, { fontSize: 16, bold: true, color: fg });
    para(s, L_SHORT, c.x + 0.423, c.y + 0.574, 2.379, 0.867, { color: fg });
  });
}

// 5 — "Our Product Digital Arena"
function slide05(p) {
  const s = p.addSlide();
  navBar(s);
  twoToneHeading(s, 'Our Product ', 'Digital Arena', 1.039, 1.315, 3.284, 1.447, { inline: true });
  para(s, L_LONG, 1.039, 3.226, 4.931, 0.603);

  ['01. Product', '02. Product', '03. Product'].forEach((label, i) => {
    const y = 1.353 + i * 1.9;
    text(s, label, 7.363, y, 2.379, 0.37, { fontSize: 16, bold: true, color: RED });
    para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed doLorem ipsum dolor sit amet, ',
      7.363, y + 0.363, 4.887, 0.601);
  });
  [1.159, 2.775, 4.392].forEach(x => imagePlaceholder(s, x, 4.54, 1.492, 1.578, { noCaption: true }));
}

// 6 — "Our Team Digital Arena"
const S6_MEMBERS = [
  { x: 0.925, name: 'Harper', role: 'Rusher', tx: 1.04, ty: 5.263, rx: 1.051, ry: 5.792, w: 0.976 },
  { x: 4.025, name: 'Alfredo', role: 'Support', tx: 4.171, ty: 5.282, rx: 4.183, ry: 5.811, w: 1.09 },
  { x: 7.125, name: 'Drew', role: 'Flanker', tx: 7.416, ty: 5.282, rx: 7.428, ry: 5.811, w: 0.873 },
  { x: 10.226, name: 'Bailey', role: 'Sniper', tx: 10.444, ty: 5.282, rx: 10.456, ry: 5.811, w: 0.873 },
];
function slide06(p) {
  const s = p.addSlide();
  navBar(s);
  S6_MEMBERS.forEach(m => {
    // flowChartInputOutput rotated 90° gives the slanted red flag behind each portrait
    shape(s, 'flowChartInputOutput', m.x, 4.161, 3.139, 0.761, { fill: { color: RED }, rotate: 90 });
    rect(s, m.x + 0.122, 3.542, 1.281, 1.281, { fill: { color: WHITE } });
    imagePlaceholder(s, m.x + 0.218, 3.638, 1.089, 1.089, { noCaption: true });
    text(s, m.name, m.tx, m.ty, m.w, 0.37, { fontSize: 16, bold: true, color: BLACK });
    text(s, m.role, m.rx, m.ry, m.w, 0.286, { fontSize: 11, color: BLACK });
  });
  twoToneHeading(s, 'Our Team', 'Digital Arena', 1.047, 1.319, 3.284, 1.447);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam tincidunt ante nec sem congue ' +
    'convallis. Pellentesque vel mauris quis nisl ornare rutrum in id',
    7.168, 1.182, 5.282, 1.035, { valign: 'middle', margin: [19.2, 19.2, 9.6, 9.6] });
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aliquam tincidunt ante nec sem congue convallis.',
    7.185, 2.011, 5.282, 0.77, { valign: 'middle', margin: [19.2, 19.2, 9.6, 9.6] });
}

// 7 — "Best Service Digital Arena" — three service cards on a red band
const S7_CARDS = [
  { x: 1.158, label: '01. Service', dark: true, icon: 'sun', ix: 2.503, iy: 3.778, iw: 0.47, ih: 0.47 },
  { x: 5.089, label: '02. Service', dark: false, icon: 'monitor', ix: 6.459, iy: 3.847, iw: 0.409, ih: 0.334 },
  { x: 9.02, label: '03. Service', dark: false, icon: 'chat', ix: 10.353, iy: 3.847, iw: 0.484, ih: 0.334 },
];
function slide07(p) {
  const s = p.addSlide();
  imagePlaceholder(s, 8.249, 0, 5.085, 2.736);
  rect(s, 3.514, 2.736, 9.819, 4.764, { fill: { color: RED } });
  rect(s, 1.263, 3.205, 3.16, 3.795, { fill: { color: WHITE } });
  navBar(s, BLACK, WHITE);

  S7_CARDS.forEach(c => {
    const fg = c.dark ? WHITE : BLACK;
    rect(s, c.x, 3.292, 3.16, 3.622, { fill: { color: c.dark ? RED : WHITE } });
    text(s, c.label, c.x + 0.391, 4.471, 2.379, 0.37, { fontSize: 16, bold: true, color: fg, align: 'center' });
    text(s, L_TEMPOR, c.x + 0.472, 5.063, 2.216, 1.132,
      { align: 'center', lineSpacingMultiple: 1.5, color: fg });
    icon(s, c.icon, c.ix, c.iy, c.iw, c.ih, c.dark ? WHITE : INK);
  });
  heading(s, [
    { text: 'Best Service ', options: { color: BLACK } },
    { text: 'Digital Arena', options: { color: RED } },
  ], 1.055, 1.305, 6.262, 0.774);
}

// 8 — "Our Amazing Portfolio"
function slide08(p) {
  const s = p.addSlide();
  [[6.677, 1.393], [9.515, 1.393], [6.677, 3.846], [9.515, 3.846]]
    .forEach(([x, y]) => imagePlaceholder(s, x, y, 2.667, 2.279));
  navBar(s);
  twoToneHeading(s, 'Our Amazing', 'Portfolio', 1.042, 1.292, 3.133, 1.447);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore ' +
    'PLACEHOLDER', 1.042, 2.978, 4.128, 1.132);

  rect(s, 1.159, 4.558, 2.458, 0.209, { fill: { color: RED }, line: { color: RED, width: 1 } });
  rect(s, 3.399, 4.558, 1.692, 0.209, { fill: { color: WHITE }, line: { color: RED, width: 1 } });
  [['70%', 1.048, 5.216, 5.579], ['30%', 3.454, 5.209, 5.571]].forEach(([label, x, y1, y2]) => {
    text(s, label, x, y1, 0.855, 0.37, { fontSize: 16, bold: true, color: RED });
    para(s, 'Lorem ipsum dolor sit amet, consectetur', x, y2, 1.743, 0.601);
  });

  rect(s, 9.771, 4.064, 2.155, 1.842, { fill: { color: RED, transparency: 25 } });
  text(s, '18+', 10.262, 4.409, 1.172, 0.774, { fontSize: 40, bold: true, color: WHITE });
  text(s, 'Years', 10.321, 5.096, 1.054, 0.37, { fontSize: 16, bold: true, color: WHITE, align: 'center' });
}

// 9 — Infographic: pyramid of three panels + three numbered columns
const S9_COLS = [
  { x: 5.325, n: '01', color: INK },
  { x: 7.898, n: '02', color: RED },
  { x: 10.471, n: '03', color: INK },
];
function slide09(p) {
  const s = p.addSlide();
  poly(s, [[26.9, 0, 1], [100, 73.1], [54.3, 99.6], [52.7, 100], [0, 100], [26.9, 0], []],
    2.226, 3.037, 1.845, 1.838, { fill: { color: RED }, rotate: 15 });
  poly(s, [[41.9, 0, 1], [100, 71.8], [15.1, 100], [0, 30], [41.9, 0], []],
    2.879, 4.665, 2.047, 1.651, { fill: { color: INK }, rotate: 15 });
  poly(s, [[25.2, 0, 1], [82.2, 0], [100, 71.3], [0, 100], [25.2, 0], []],
    1.331, 4.734, 1.738, 1.622, { fill: { color: INK }, rotate: 15 });
  rect(s, 2.474, 4.318, 1.0, 1.0, { fill: { color: WHITE } });
  icon(s, 'barchart', 2.737, 4.601, 0.475, 0.434, INK);
  icon(s, 'envelope', 2.767, 3.702, 0.407, 0.254, WHITE);
  icon(s, 'microscope', 3.697, 5.308, 0.417, 0.474, WHITE);
  icon(s, 'sitemap', 1.8, 5.327, 0.455, 0.455, WHITE);

  S9_COLS.forEach(c => {
    text(s, c.n, c.x, 2.932, 0.758, 0.64, { fontSize: 32, bold: true, color: c.color });
    text(s, 'Title Here', c.x, 3.538, 1.226, 0.337, { fontSize: 14, bold: true, color: c.color });
    text(s, L_ADIPI, c.x, 3.841, 1.852, 1.182, { fontSize: 11, lineSpacingMultiple: 1.5 });
  });
  rect(s, 5.884, 5.308, 6.294, 0.803, { fill: { color: RED } });
  text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. ' +
    'Fusce posuere, magna Fusce posuere,', 6.015, 5.408, 6.032, 0.603,
    { align: 'center', lineSpacingMultiple: 1.5, color: WHITE });
  sectionTitle(s);
  navBar(s);
}

// 10 — Infographic: three speech-bubble cards
const S10_CARDS = [
  { x: 1.143, color: RED, icon: 'microscope', ix: 3.319, iy: 4.836, iw: 0.613, ih: 0.696 },
  { x: 5.071, color: INK, icon: 'sitemap', ix: 7.281, iy: 4.889, iw: 0.591, ih: 0.591 },
  { x: 9.0, color: RED, icon: 'trophy', ix: 11.123, iy: 4.836, iw: 0.753, ih: 0.696 },
];
function slide10(p) {
  const s = p.addSlide();
  S10_CARDS.forEach(c => {
    rect(s, c.x, 2.976, 3.19, 0.619, { fill: { color: c.color } });
    shape(s, 'wedgeRectCallout', c.x, 4.25, 3.19, 1.869, { fill: { color: c.color }, flipV: true });
    text(s, L_ADIPI2, c.x + 0.163, 4.455, 1.647, 1.458, { fontSize: 11, lineSpacingMultiple: 1.5, color: WHITE });
    rect(s, c.x + 2.0, 4.455, 1.0, 1.46, { fill: { color: WHITE } });
    text(s, ' YourTitle Here', c.x + 0.503, 3.084, 2.183, 0.404,
      { fontSize: 18, bold: true, color: WHITE, align: 'center' });
    icon(s, c.icon, c.ix, c.iy, c.iw, c.ih, c.color);
  });
  sectionTitle(s);
  navBar(s);
}

// Three preset arrows in the deck carry custom adjust handles that pptxgenjs cannot
// emit, so they are drawn as freeform polygons with the reference's proportions.
const ARROW_RIGHT = [[0, 21.4, 1], [80.8, 21.4], [80.8, 0], [100, 50], [80.8, 100], [80.8, 78.6], [0, 78.6], []];
const ARROW_UP = [[0, 17.7, 1], [50, 0], [100, 17.7], [73.6, 17.7], [73.6, 100], [26.4, 100], [26.4, 17.7], []];
const ARROW_BENT_UP = [[0, 74.4, 1], [72.3, 74.4], [72.3, 28.5], [60.3, 28.5], [79.4, 0],
  [98.5, 28.5], [86.5, 28.5], [86.5, 100], [0, 100], []];

// 11 — Infographic: big up-arrow with two outlined bars
const S11_ROWS = [
  { y: 3.359, color: RED_MD, iconColor: RED, icon: 'building', iy: 3.713, ih: 0.423 },
  { y: 4.942, color: INK, iconColor: INK, icon: 'person', iy: 5.298, ih: 0.42 },
];
function slide11(p) {
  const s = p.addSlide();
  poly(s, ARROW_UP, 1.179, 2.216, 1.679, 5.284, { fill: { color: RED } });
  S11_ROWS.forEach((r, i) => {
    rect(s, 1.452, r.y, 1.131, 1.131, { fill: { color: r.color } });
    rect(s, 1.571, r.y + 0.119, 0.893, 0.893, { fill: { color: WHITE } });
    rect(s, 2.583, r.y, 9.571, 1.131, { fill: { color: r.color } });
    rect(s, 2.583, r.y + 0.129, 9.448, 0.873, { fill: { color: WHITE } });
    icon(s, r.icon, 1.809, r.iy, 0.417, r.ih, r.iconColor);
    text(s, L_MAECENAS, 2.702, r.y + 0.231, 9.179, 0.627,
      { fontSize: 11, align: 'justify', lineSpacingMultiple: 1.5 });
  });
  sectionTitle(s);
  navBar(s);
}

// 12 — Infographic: four-petal pinwheel with diamond icon badges
const S12_PETALS = [
  { x: 5.273, y: 4.383, rot: 0, color: INK },
  { x: 6.336, y: 4.383, rot: -90, color: RED_LT },
  { x: 6.34, y: 3.317, rot: 180, color: INK_LT },
  { x: 5.273, y: 3.32, rot: 90, color: RED },
];
const S12_BADGES = [
  { x: 4.805, y: 2.956, color: RED, icon: 'bookClosed', ix: 5.232, iy: 3.336, iw: 0.37, ih: 0.464 },
  { x: 7.303, y: 2.965, color: INK, icon: 'bookOpen', ix: 7.685, iy: 3.394, iw: 0.462, ih: 0.366 },
  { x: 4.807, y: 5.345, color: INK, icon: 'clipboard', ix: 5.211, iy: 5.699, iw: 0.416, ih: 0.517 },
  { x: 7.259, y: 5.345, color: RED, icon: 'house', ix: 7.631, iy: 5.737, iw: 0.48, ih: 0.441 },
];
const S12_OPTIONS = [
  { label: 'Option 01', color: RED, x: 1.114, y: 2.941, tx: 1.125, ty: 3.4 },
  { label: 'Option 02', color: INK, x: 9.843, y: 2.888, tx: 9.854, ty: 3.347 },
  { label: 'Option 03', color: INK, x: 1.114, y: 4.805, tx: 1.125, ty: 5.264 },
  { label: 'Option 04', color: RED, x: 9.843, y: 4.752, tx: 9.854, ty: 5.211 },
];
const PETAL = [[0.4, 0, 1], [100, 100, 55.4, 0, 100, 44.8], [57.2, 100],
  [0.4, 42.9, 57.2, 68.5, 31.8, 42.9], [0, 43], [0, 0], [0.4, 0], []];
function slide12(p) {
  const s = p.addSlide();
  S12_PETALS.forEach(pt => poly(s, PETAL, pt.x, pt.y, 1.852, 1.845, { fill: { color: pt.color }, rotate: pt.rot }));
  S12_BADGES.forEach(b => {
    shape(s, 'diamond', b.x, b.y, 1.225, 1.225, { fill: { color: b.color } });
    icon(s, b.icon, b.ix, b.iy, b.iw, b.ih, WHITE);
  });
  S12_OPTIONS.forEach(o => {
    text(s, o.label, o.x, o.y, 1.946, 0.505, { fontSize: 24, bold: true, color: o.color });
    text(s, L_ADIPI2, o.tx, o.ty, 2.366, 0.904, { fontSize: 11, lineSpacingMultiple: 1.5 });
  });
  sectionTitle(s);
  navBar(s);
}

// 13 — Infographic: four right-arrows with diamond numbers
const S13_ARROWS = [
  { x: 1.167, y: 2.972, n: '01', color: RED, ring: RED_DK },
  { x: 7.449, y: 2.971, n: '02', color: INK, ring: INK_DK },
  { x: 1.167, y: 4.896, n: '03', color: INK, ring: INK_DK },
  { x: 7.449, y: 4.894, n: '04', color: RED, ring: RED_DK },
];
function slide13(p) {
  const s = p.addSlide();
  S13_ARROWS.forEach(a => {
    poly(s, ARROW_RIGHT, a.x, a.y, 4.718, 1.556, { fill: { color: a.color } });
    shape(s, 'diamond', a.x + 0.023, a.y - 0.027, 1.61, 1.61, { fill: { color: a.ring } });
    shape(s, 'diamond', a.x + 0.247, a.y + 0.197, 1.163, 1.163, { fill: { color: WHITE } });
    text(s, a.n, a.x + 0.247, a.y + 0.197, 1.163, 1.163,
      { fontSize: 16, bold: true, color: a.color, align: 'center', valign: 'middle' });
    text(s, 'Title Here', a.x + 1.706, a.y + 0.474, 1.212, 0.303, { fontSize: 12, bold: true, color: WHITE });
    text(s, 'Lorem ipsum dolor sit amet, ', a.x + 1.706, a.y + 0.699, 2.389, 0.349,
      { fontSize: 11, color: WHITE, lineSpacingMultiple: 1.5 });
  });
  sectionTitle(s);
  navBar(s);
}

// 14 — Infographic: five snipped-corner tags
const S14_TAGS = [
  { x: 1.171, n: '01', color: RED, icon: 'dollar', ix: 1.807, iw: 0.409 },
  { x: 3.499, n: '02', color: INK, icon: 'globe', ix: 4.116, iw: 0.427 },
  { x: 5.826, n: '03', color: RED, icon: 'yen', ix: 6.455, iw: 0.409 },
  { x: 8.154, n: '04', color: INK, icon: 'euro', ix: 8.775, iw: 0.425 },
  { x: 10.482, n: '05', color: RED, icon: 'gem', ix: 11.124, iw: 0.396 },
];
function slide14(p) {
  const s = p.addSlide();
  S14_TAGS.forEach(t => {
    shape(s, 'snip2SameRect', t.x, 2.972, 1.681, 3.139, { fill: { color: t.color } });
    text(s, 'Title Here', t.x + 0.189, 3.089, 1.289, 0.337,
      { fontSize: 14, bold: true, color: WHITE, align: 'center' });
    text(s, 'Lorem ipsum dolor sit amet, Lorem', t.x + 0.203, 3.475, 1.275, 0.904,
      { fontSize: 11, color: WHITE, align: 'center', lineSpacingMultiple: 1.5 });
    shape(s, 'snip2SameRect', t.x + 0.434, 4.51, 0.815, 0.815, { fill: { color: WHITE } });
    text(s, t.n, t.x + 0.434, 4.51, 0.815, 0.815,
      { fontSize: 24, bold: true, color: t.color, align: 'center', valign: 'middle' });
    icon(s, t.icon, t.ix, 5.512, t.iw, 0.409, WHITE);
  });
  sectionTitle(s);
  navBar(s);
}

// 15 — Infographic: staircase of three numbered blocks + step copy
const S15_BLOCKS = [
  { x: 0.585, y: 3.2, n: '01', color: INK, ax: 0.0, ay: 5.014, flipV: false, acolor: INK },
  { x: 2.705, y: 4.586, n: '02', color: RED, ax: 2.121, ay: 3.2, flipV: true, acolor: RED },
  { x: 4.826, y: 3.2, n: '03', color: INK, ax: 4.241, ay: 5.014, flipV: false, acolor: INK },
];
function slide15(p) {
  const s = p.addSlide();
  S15_BLOCKS.forEach(b => {
    rect(s, b.x, b.y, 1.305, 1.305, { fill: { color: b.color } });
    text(s, b.n, b.x, b.y, 1.305, 1.305,
      { fontSize: 40, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    poly(s, ARROW_BENT_UP, b.ax, b.ay, 1.509, 0.877, { fill: { color: b.acolor }, flipV: b.flipV });
  });
  ['01. Step', '02. Step', '03. Step'].forEach((label, i) => {
    const y = 2.907 + i * 1.2205;
    text(s, label, 7.352, y, 1.18, 0.303, { fontSize: 12, bold: true, color: i === 1 ? RED : INK });
    para(s, L_STEP, 7.352, y + 0.234, 4.903, 0.603);
  });
  sectionTitle(s);
  navBar(s);
}

// 16 — Infographic: four framed squares chained with arrows
const S16_STEPS = [
  { x: 1.2, color: RED }, { x: 4.117, color: INK }, { x: 7.033, color: RED }, { x: 9.95, color: INK },
];
function slide16(p) {
  const s = p.addSlide();
  S16_STEPS.forEach((st, i) => {
    rect(s, st.x + 0.126, 3.088, 2.108, 2.108, { fill: { color: st.color } });
    rect(s, st.x, 2.962, 2.36, 2.36, { fill: NONE, line: { color: st.color, width: 1 } });
    icon(s, 'docChart', st.x + 0.925, 3.343, 0.509, 0.579, WHITE);
    text(s, ' Title Here', st.x + 0.44, 4.125, 1.481, 0.303,
      { fontSize: 12, bold: true, color: WHITE, align: 'center' });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer.', st.x + 0.126, 4.338, 2.108, 0.601,
      { align: 'center', lineSpacingMultiple: 1.5, color: WHITE });
    text(s, 'Lorem ipsum dolor sit amet, consectetuer adipiscing.', st.x - 0.024, 5.459, 2.384, 0.603,
      { align: 'center', lineSpacingMultiple: 1.5 });
    if (i < 3) shape(s, 'rightArrow', st.x + 2.369, 3.893, 0.544, 0.497, { fill: { color: st.color } });
  });
  sectionTitle(s);
  navBar(s);
}

// 17 — Infographic: six leaning columns
const S17_COLS = [
  { x: 1.23, color: RED, icon: 'palette', ix: 1.758, iy: 3.393, iw: 0.416, ih: 0.393 },
  { x: 3.118, color: INK, icon: 'gift', ix: 3.673, iy: 3.348, iw: 0.361, ih: 0.388 },
  { x: 5.006, color: RED, icon: 'palette', ix: 5.537, iy: 3.393, iw: 0.416, ih: 0.393 },
  { x: 6.894, color: INK, icon: 'gift', ix: 7.452, iy: 3.348, iw: 0.361, ih: 0.388 },
  { x: 8.782, color: RED, icon: 'palette', ix: 9.306, iy: 3.357, iw: 0.416, ih: 0.393 },
  { x: 10.644, color: INK, icon: 'gift', ix: 11.199, iy: 3.343, iw: 0.361, ih: 0.388 },
];
function slide17(p) {
  const s = p.addSlide();
  // The slanted "shadow" wedges sit behind every column except the last.
  S17_COLS.forEach((c, i) => {
    rect(s, c.x, 2.962, 1.471, 3.146, { fill: { color: c.color } });
    if (i < 5) {
      shape(s, 'flowChartInputOutput', c.x + 0.107, 4.326, 3.146, 0.417,
        { fill: { color: c.color === RED ? RED_LT : INK_LT }, rotate: 270 });
    }
  });
  S17_COLS.forEach(c => {
    icon(s, c.icon, c.ix, c.iy, c.iw, c.ih, WHITE);
    text(s, L_CEASK, c.x + 0.144, 4.218, 1.185, 1.663,
      { align: 'justify', lineSpacingMultiple: 1.5, color: WHITE });
  });
  sectionTitle(s);
  navBar(s);
}

// 18 — Infographic: four flag posts
const S18_FLAGS = [
  { x: 1.154, n: '01', color: RED, icon: 'palette', ix: 1.839, iw: 0.416, ih: 0.393 },
  { x: 4.266, n: '02', color: INK, icon: 'gift', ix: 4.978, iw: 0.361, ih: 0.388 },
  { x: 7.378, n: '03', color: RED, icon: 'palette', ix: 8.036, iw: 0.416, ih: 0.393 },
  { x: 10.49, n: '04', color: INK, icon: 'gift', ix: 11.175, iw: 0.361, ih: 0.388 },
];
function slide18(p) {
  const s = p.addSlide();
  S18_FLAGS.forEach(f => {
    rect(s, f.x + 0.096, 2.962, 1.594, 1.594, { fill: { color: f.color } });
    rect(s, f.x, 2.962, 0.096, 2.346, { fill: { color: POLE } }); // vertical grey pole
    text(s, f.n, f.x + 0.427, 3.11, 0.931, 0.572,
      { fontSize: 28, bold: true, color: WHITE, align: 'center' });
    icon(s, f.icon, f.ix, 3.862, f.iw, f.ih, WHITE);
    para(s, 'Lorem ipsum dolor sit consectetuer amet, consectetuer adipiscing ceask',
      f.x + 0.064, 5.064, 1.594, 1.133);
  });
  sectionTitle(s);
  navBar(s);
}

// 19 — Contact page
const S19_CONTACT = [
  { icon: 'phone', ix: 6.361, iy: 4.018, iw: 0.418, ih: 0.393, ty: 4.03, label: '+12345678910', size: 12 },
  { icon: 'envelope', ix: 6.361, iy: 4.887, iw: 0.416, ih: 0.259, ty: 4.831, label: 'www.yourmail.com', size: 12 },
  { icon: 'building', ix: 6.361, iy: 5.71, iw: 0.418, ih: 0.423, ty: 5.707, label: 'East 87 Street', size: 10.5 },
];
const S19_CONTACT2 = [
  { icon: 'chat', ix: 9.807, iy: 4.085, iw: 0.413, ih: 0.329, ty: 4.03, label: '+12345678910', size: 12 },
  { icon: 'qr', ix: 9.85, iy: 4.877, iw: 0.326, ih: 0.323, ty: 4.831, label: 'https.yourweb', size: 12 },
  { icon: 'notebook', ix: 9.818, iy: 5.747, iw: 0.39, ih: 0.365, ty: 5.707, label: 'Los Angeles, LA 10067', size: 10.5 },
];
function slide19(p) {
  const s = p.addSlide();
  poly(s, [[0.1, 0, 1], [13, 0], [100, 92.2], [93.3, 100], [87, 100], [0, 7.8], [0, 0.1], []],
    0, 0, 4.611, 4.136, { fill: { color: INK } });
  navBar(s, BLACK);
  icon(s, 'gamepad', 1.159, 0.413, 0.444, 0.236, RED);
  rect(s, 1.082, 5.335, 1.991, 0.57, { fill: { color: RED }, rotate: 43.5 });
  poly(s, [[0, 0, 1], [100, 73.4], [68.5, 100], [0, 49.7], []],
    0, 0.04, 4.914, 6.241, { fill: { color: PHOTO } });
  text(s, '[image]', 0.6, 3.0, 2.0, 0.36, { color: PHOTO_TXT, fontSize: 12, align: 'center' });

  twoToneHeading(s, 'If Need Any Info', 'Please Contact Us', 6.251, 1.318, 4.225, 1.447);
  para(s, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt Lorem ' +
    'Lorem ipsum dolor sit amet', 6.251, 2.835, 6.043, 0.601);

  [[S19_CONTACT, 7.077], [S19_CONTACT2, 10.519]].forEach(([rows, tx]) => {
    rows.forEach(r => {
      icon(s, r.icon, r.ix, r.iy, r.iw, r.ih, INK);
      text(s, r.label, tx, r.ty, 1.867, 0.37, { fontSize: r.size, lineSpacingMultiple: 1.5 });
    });
  });
}

// 20 — "Thank You"
function slide20(p) {
  const s = p.addSlide();
  s.background = { color: PHOTO };
  rect(s, 0, 0, 13.333, 7.5, { fill: { color: RED, transparency: 25 } });
  navBar(s, WHITE);
  rect(s, 4.116, 1.594, 5.101, 3.722, { fill: NONE, line: { color: WHITE, width: 1 } });
  text(s, 'Thank You', 4.116, 1.594, 5.101, 3.722,
    { fontFace: HEAD, fontSize: 96, color: WHITE, align: 'center', valign: 'middle' });
  text(s, 'Esport Presentation Template', 4.707, 5.765, 3.918, 0.404,
    { fontSize: 18, color: WHITE, align: 'center', wrap: false });
}

// ── Build ──────────────────────────────────────────────────────────────────
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE';
  pptx.author = 'Digital Arena';
  pptx.title = 'Digital Arena — Esport Presentation Template';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
   slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach(fn => fn(pptx));

  return pptx.writeFile({ fileName: path.join(__dirname, '0fe37884-c2fe-4af2-9a8b-2002fc25ef32_grok_final.pptx') });
}

build().then(f => console.log('wrote', f)).catch(err => { console.error(err); process.exit(1); });
