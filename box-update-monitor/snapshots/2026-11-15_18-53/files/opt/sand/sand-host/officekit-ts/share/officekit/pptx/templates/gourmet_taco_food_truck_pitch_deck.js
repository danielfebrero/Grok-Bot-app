/*
 * Delizioso Pitch Deck - recreated with pptxgenjs.
 * Run: node 17b26c2b-a18c-4ec9-ab35-f456ad1e2688_grok_final.js
 */
'use strict';

const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------- design tokens
const YELLOW = 'F5D147';
const BLACK = '000000';
const WHITE = 'FFFFFF';
const GRAY = '808080';
const DEVICE = '2C2E30';
const FONT = 'Plus Jakarta Sans';
const LS = 1.5; // the deck sets 150% line spacing on every body paragraph

const HEADLINE = 'Redefining Gourmet Tacos And Burritos';
const TRUCK = 'Mobile Food Truck';
const SITE = 'Delizioso.website';

// ---------------------------------------------------------------- primitives
/** Solid colour block (the yellow panels that come from the slide layouts). */
function panel(slide, x, y, w, h, fill) {
  slide.addShape('rect', { x, y, w, h, fill: { color: fill || YELLOW }, line: { type: 'none' } });
}

/** Thin 2pt rule used as a section divider on nearly every slide. */
function rule(slide, x, y, w, color) {
  slide.addShape('line', { x, y, w, h: 0, line: { color: color || BLACK, width: 2 } });
}

/** Text box. Reference deck leaves the OOXML insets at their defaults, so do we. */
function text(slide, content, o) {
  slide.addText(content, Object.assign({
    fontFace: FONT, fontSize: 12, color: BLACK,
    align: 'left', valign: 'top', wrap: false, fit: 'resize',
  }, o));
}

/**
 * Black pill button with a white centred caption.
 * `cap` mirrors the reference's separate caption text box: {x, y, w, size}.
 */
function pill(slide, x, y, w, h, label, cap) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: h / 2, fill: { color: BLACK }, line: { type: 'none' },
  });
  text(slide, label, {
    x: cap.x, y: cap.y, w: cap.w, h: cap.size > 12 ? 0.37 : 0.303,
    fontSize: cap.size || 12, color: WHITE, align: 'center',
  });
}

/** Brand mark: outlined white carrot on a black disc (replaces the icon graphic). */
function logo(slide, x, y, d) {
  const stroke = { color: WHITE, width: Math.max(0.75, 1.6 * d) };
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: BLACK }, line: { type: 'none' } });
  slide.addShape('teardrop', {
    x: x + 0.14 * d, y: y + 0.24 * d, w: 0.58 * d, h: 0.58 * d,
    rotate: 225, fill: { type: 'none' }, line: stroke,
  });
  // leaf sprigs fanning off the carrot shoulder
  slide.addShape('line', { x: x + 0.66 * d, y: y + 0.14 * d, w: 0.02 * d, h: 0.16 * d, line: stroke });
  slide.addShape('line', { x: x + 0.66 * d, y: y + 0.22 * d, w: 0.17 * d, h: 0.02 * d, line: stroke });
  slide.addShape('line', { x: x + 0.68 * d, y: y + 0.14 * d, w: 0.13 * d, h: 0.13 * d, flipV: true, line: stroke });
  // rib marks crossing the carrot body
  slide.addShape('line', { x: x + 0.30 * d, y: y + 0.54 * d, w: 0.12 * d, h: 0.12 * d, line: stroke });
  slide.addShape('line', { x: x + 0.44 * d, y: y + 0.40 * d, w: 0.12 * d, h: 0.12 * d, line: stroke });
}

/** "Next" affordance: outlined disc holding a chevron (replaces the icon graphic). */
function arrowDisc(slide, x, y, d) {
  const c = d / 2;
  const a = 0.24 * d;
  const stroke = { color: BLACK, width: 2 };
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { type: 'none' }, line: stroke });
  slide.addShape('line', { x: x + c - 0.6 * a, y: y + c - a, w: a, h: a, line: stroke });
  slide.addShape('line', { x: x + c - 0.6 * a, y: y + c, w: a, h: a, flipV: true, line: stroke });
}

/**
 * Stand-in for the device photographs: a dark body with a coloured screen.
 * `side`/`chin` are the bezel thicknesses measured off the reference render.
 */
function deviceShot(slide, x, y, w, h, screen, radius, side, chin, notch) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: radius, fill: { color: DEVICE }, line: { type: 'none' },
  });
  slide.addShape('roundRect', {
    x: x + side, y: y + chin, w: w - 2 * side, h: h - 2 * chin,
    rectRadius: Math.max(radius - side, 0.02), fill: { color: screen }, line: { type: 'none' },
  });
  if (notch) {
    slide.addShape('roundRect', {
      x: x + 0.26 * w, y: y + chin - 0.02, w: 0.48 * w, h: 0.055 * h,
      rectRadius: 0.06, fill: { color: DEVICE }, line: { type: 'none' },
    });
  }
  text(slide, '[image]', {
    x, y: y + h / 2 - 0.15, w, h: 0.3,
    fontSize: 10, color: GRAY, align: 'center', wrap: true,
  });
}

/** Logo lock-up in the top-left plus the optional right-hand strap line. */
function header(slide, right) {
  logo(slide, 0.67, 0.31, 0.522);
  text(slide, 'Delizioso', { x: 1.287, y: 0.42, w: 0.928, h: 0.303 });
  if (right === TRUCK) text(slide, TRUCK, { x: 11.01, y: 0.42, w: 1.653, h: 0.303, align: 'right' });
  if (right === HEADLINE) text(slide, HEADLINE, { x: 9.325, y: 0.42, w: 3.338, h: 0.303, align: 'right' });
}

// Shorthand builders for the three recurring paragraph styles.
const lead = (o) => Object.assign({ fontSize: 10, color: GRAY, wrap: true, lineSpacingMultiple: LS }, o);
const body = (o) => Object.assign({ fontSize: 10, color: BLACK, wrap: true, lineSpacingMultiple: LS }, o);
const note = (o) => Object.assign({ fontSize: 12, color: BLACK, wrap: true, lineSpacingMultiple: LS }, o);

// ---------------------------------------------------------------- slides
function slide01(pptx) {
  const s = pptx.addSlide();
  s.background = { color: YELLOW };
  text(s, [
    { text: 'Delizioso', options: { breakLine: true } },
    { text: 'Pitch Deck ' },
  ], { x: 1.843, y: 1.584, w: 6.872, h: 3.063, fontSize: 88, bold: true, valign: 'middle' });
  pill(s, 1.912, 5.128, 2.855, 0.787, 'Get Started', { x: 2.608, y: 5.337, w: 1.462, size: 16 });
  text(s, 'Mobile Food Truck ', { x: 5.279, y: 5.354, w: 1.943, h: 0.337, fontSize: 14, align: 'center' });
  header(s, HEADLINE);
  arrowDisc(s, 12.149, 6.653, 0.492);
  rule(s, 0.67, 6.899, 10.722, WHITE);
}

function slide02(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 2.829, 7.5);
  s.addShape('round2SameRect', {
    x: 7.181, y: -0.23, w: 3.49, h: 6.998, rotate: 270,
    fill: { color: YELLOW }, line: { type: 'none' },
  });
  logo(s, 5.978, 1.922, 0.64);
  text(s, 'Introduction', { x: 5.978, y: 2.842, w: 5.84, h: 1.212, fontSize: 66 });
  text(s, 'Delizioso is a mobile food truck specializing in gourmet tacos and burritos. ',
    body({ x: 5.978, y: 4.17, w: 5.169, h: 0.32 }));
  text(s, 'PLACEHOLDER' +
    'Mexican flavors with innovative and high-quality ingredients. ',
    lead({ x: 3.96, y: 5.895, w: 3.721, h: 0.825 }));
  text(s, [
    { text: 'About', options: { breakLine: true } },
    { text: 'Delizioso' },
  ], { x: 0.65, y: 6.022, w: 1.049, h: 0.572, fontSize: 14 });
  header(s, TRUCK);
  arrowDisc(s, 12.18, 6.062, 0.492);
  rule(s, 9.045, 6.308, 2.377);
}

function slide03(pptx) {
  const s = pptx.addSlide();
  panel(s, 8.196, 0, 5.138, 7.5);
  deviceShot(s, 9.513, 1.154, 2.502, 5.192, YELLOW, 0.28, 0.16, 0.17, true);
  text(s, 'Our Mission', { x: 1.154, y: 1.536, w: 4.514, h: 1.01, fontSize: 54, valign: 'middle' });
  arrowDisc(s, 6.891, 2.766, 0.492);
  rule(s, 1.253, 3.012, 4.882);
  logo(s, 1.253, 3.784, 0.522);
  text(s, 'Delizioso', { x: 1.87, y: 3.894, w: 0.928, h: 0.303 });
  text(s, 'Our mission is to provide delicious, memorable meals using carefully selected ingredients ' +
    'and expertly crafted recipes. Our passionate chefs constantly strive to innovate and create ' +
    'exciting new culinary creations.',
    lead({ x: 1.253, y: 4.634, w: 3.396, h: 1.33 }));
  header(s, null);
}

function slide04(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 3.996, 13.333, 3.504);
  arrowDisc(s, 12.149, 6.353, 0.492);
  rule(s, 0.67, 6.599, 10.722, WHITE);
  deviceShot(s, 4.956, 3.13, 3.421, 5.233, YELLOW, 0.16, 0.22, 0.55);
  text(s, 'About Our Vision', { x: 3.8, y: 1.787, w: 5.724, h: 0.909, fontSize: 48, align: 'center' });
  pill(s, 5.235, 0.9, 2.855, 0.522, SITE, { x: 5.853, y: 1.01, w: 1.618, size: 12 });
  text(s, 'PLACEHOLDER',
    { x: 1.563, y: 4.777, w: 2.708, h: 0.97, align: 'right', wrap: true, lineSpacingMultiple: 1.5 });
  text(s, 'PLACEHOLDER' +
    'customers coming back for more',
    body({ x: 8.881, y: 4.849, w: 3.396, h: 0.825 }));
  header(s, TRUCK);
}

function slide05(pptx) {
  const s = pptx.addSlide();
  panel(s, 0.302, 1.191, 0.891, 5.118);
  text(s, 'Signature Tacos', { x: 3.407, y: 1.664, w: 1.515, h: 0.303 });
  text(s, 'Enjoy our delicious signature tacos, such as Braised Short Rib with Chipotle Aioli, ' +
    'Grilled Shrimp with Mango Salsa, and Vegetarian Black Bean with Avocado Crema.',
    lead({ x: 3.407, y: 2.024, w: 3.063, h: 1.077 }));
  text(s, 'Creative Burritos', { x: 3.369, y: 4.415, w: 1.587, h: 0.303 });
  text(s, 'Creative Burritos: Our unique burritos that mix classic Mexican flavors with fun ' +
    'surprises. Taste our Grilled Chicken and Pineapple Burrito with Cilantro-Lime Rice',
    lead({ x: 3.369, y: 4.776, w: 3.101, h: 1.077 }));
  text(s, [
    { text: 'Delizioso', options: { breakLine: true } },
    { text: 'Menu' },
  ], { x: 7.347, y: 1.491, w: 4.195, h: 2.322, fontSize: 66 });
  panel(s, 7.348, 4.67, 5.985, 1.646);
  text(s, 'Our Menu Features A Tantalizing Selection Of Gourmet Tacos And Burritos, Each Bursting ' +
    'With Unique Flavors And Made With Care. ',
    note({ x: 8.276, y: 5.008, w: 4.13, h: 0.97 }));
  header(s, null);
  arrowDisc(s, 12.171, 0.955, 0.492);
  rule(s, 7.347, 1.201, 4.067);
}

function slide06(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 4.606, 6.765, 2.913);
  text(s, 'Signature Tacos', { x: 1.314, y: 1.748, w: 4.472, h: 0.774, fontSize: 40, valign: 'middle' });
  text(s, 'PLACEHOLDER' +
    'Chipotle Aioli, Grilled Shrimp with Mango Salsa, and Vegetarian Black Bean with Avocado Crema.',
    lead({ x: 1.314, y: 2.789, w: 4.483, h: 0.825 }));
  text(s, 'Our Tacos Are Made With Fresh, High-quality Ingredients For An Unforgettable Culinary ' +
    'Experience.',
    note({ x: 1.332, y: 5.732, w: 4.509, h: 0.667 }));
  logo(s, 6.429, 5.727, 0.673);
  header(s, null);
  arrowDisc(s, 12.18, 0.835, 0.492);
  rule(s, 6.765, 1.081, 4.658);
}

function slide07(pptx) {
  const s = pptx.addSlide();
  panel(s, 10.363, 2.869, 2.97, 4.631);
  text(s, 'Creative Burritos', { x: 1.211, y: 1.396, w: 5.554, h: 0.909, fontSize: 48, valign: 'middle' });
  pill(s, 1.215, 2.867, 2.067, 0.522, SITE, { x: 1.44, y: 2.976, w: 1.618, size: 12 });
  text(s, 'Each burrito is made with care and features the freshest ingredients, ensuring a ' +
    'delicious and satisfying meal every time.',
    lead({ x: 4.161, y: 2.858, w: 3.755, h: 0.825 }));
  text(s, 'Grilled Chicken and Pineapple Burrito with Cilantro-Lime Rice: A sweet and tangy flavor ' +
    'explosion.',
    lead({ x: 2.808, y: 5.194, w: 2.229, h: 1.077 }));
  text(s, 'Spicy Cauliflower and Quinoa Burrito with Chipotle Mayo: A bold and zesty kick.',
    lead({ x: 5.597, y: 5.194, w: 2.229, h: 0.825 }));
  header(s, HEADLINE);
  arrowDisc(s, 7.425, 4.109, 0.492);
  rule(s, 0.661, 4.355, 5.847);
}

function slide08(pptx) {
  const s = pptx.addSlide();
  panel(s, 12.179, 0, 1.154, 7.5);
  panel(s, 0, 1.191, 2.251, 1.71);
  const reasons = [
    { n: '1. ', title: 'Quality Ingredients', x: 0.661, y: 3.36, tw: 1.876,
      copy: 'We believe that exceptional food starts with exceptional ingredients. At Delizioso, we ' +
        'source the freshest produce, premium meats, and authentic Mexican spices to ensure that ' +
        'every bite is a burst of flavor.' },
    { n: '3. ', title: 'Innovative Flavors', x: 6.371, y: 3.36, tw: 1.825,
      copy: 'PLACEHOLDER' +
        'unexpected flavors. Our fusion of traditional and modern culinary techniques creates a ' +
        'unique and unforgettable taste experience.' },
    { n: '2. ', title: 'Mobile Convenience', x: 0.661, y: 5.004, tw: 2.02,
      copy: "Our mobile food truck brings our gourmet tacos and burritos directly to our customers' " +
        "doorsteps. Whether it's a quick lunch break or a special event, Delizioso is there to " +
        'provide a convenient and delicious dining option.' },
    { n: '4. ', title: 'Passionate Team', x: 6.371, y: 5.004, tw: 1.759,
      copy: 'Our team is dedicated to delivering an exceptional customer experience. From our ' +
        'friendly and knowledgeable staff to our skilled chefs, everyone at Delizioso shares a ' +
        'passion for food and a commitment to excellence.' },
  ];
  reasons.forEach((r) => {
    text(s, [
      { text: r.n, options: { bold: true } },
      { text: r.title },
    ], { x: r.x, y: r.y, w: r.tw, h: 0.303 });
    text(s, r.copy, lead({ x: r.x, y: r.y + 0.361, w: 5.316, h: 0.825 }));
  });
  text(s, 'Why Choose Us?', { x: 1.239, y: 1.469, w: 4.786, h: 0.774, fontSize: 40 });
  text(s, HEADLINE, { x: 1.239, y: 2.342, w: 3.338, h: 0.303, align: 'right' });
  header(s, null);
  arrowDisc(s, 11.905, 6.653, 0.492);
  rule(s, 0.661, 6.899, 10.327);
}

function slide09(pptx) {
  const s = pptx.addSlide();
  panel(s, 4.993, 2.985, 6.017, 1.18);
  text(s, 'Our Food Truck Experience', { x: 2.437, y: 1.191, w: 7.367, h: 0.774, fontSize: 40 });
  text(s, 'Delizioso offers more than just great food; we provide an immersive dining experience. ' +
    "When you visit our food truck, you'll be greeted by vibrant colors, inviting aromas, and a " +
    'lively atmosphere. We take pride in creating a welcoming space where our customers can enjoy ' +
    'their meals and connect with friends, family, and fellow food enthusiasts.',
    lead({ x: 6.765, y: 4.836, w: 5.414, h: 1.33 }));
  text(s, 'Redefining Gourmet Tacos and Burritos',
    { x: -0.733, y: 5.086, w: 3.324, h: 0.303, align: 'right', rotate: 270 });
  header(s, TRUCK);
  arrowDisc(s, 9.803, 3.329, 0.492);
  rule(s, 6.667, 3.575, 2.313, WHITE);
}

function slide10(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 4.824, 13.333, 2.676);
  text(s, [
    { text: 'Our Commitment', options: { breakLine: true } },
    { text: 'to Quality' },
  ], { x: 1.045, y: 1.363, w: 4.83, h: 1.313, fontSize: 36, wrap: true });
  text(s, 'Quality is our top priority. We strive for excellence in all aspects of our business.',
    lead({ x: 1.098, y: 2.827, w: 4.535, h: 0.572 }));
  const pillars = [
    { n: '1.', title: 'Food Safety Standards', hx: 1.514, hw: 2.048, bx: 0.767,
      copy: 'At our establishment, we prioritize the safety and health of our customers by ' +
        'following rigorous standards for food safety and hygiene.' },
    { n: '2.', title: 'Fresh Ingredients', hx: 5.875, hw: 1.583, bx: 4.895,
      copy: 'we make it a priority to use fresh and sustainable ingredients. We work with local ' +
        'suppliers to ensure that our ingredients are of the highest quality.' },
    { n: '3.', title: 'Skilled Culinary Team', hx: 9.859, hw: 1.873, bx: 9.023,
      copy: 'PLACEHOLDER' +
        'dedicated to creating exceptional dishes.' },
  ];
  pillars.forEach((p) => {
    text(s, [
      { text: p.n, options: { bold: true, breakLine: true } },
      { text: p.title },
    ], { x: p.hx, y: 5.235, w: p.hw, h: 0.667, align: 'center', lineSpacingMultiple: 1.5 });
    text(s, p.copy, body({ x: p.bx, y: 6.013, w: 3.544, h: 0.825, align: 'center' }));
  });
  header(s, TRUCK);
  arrowDisc(s, 5.141, 3.596, 0.492);
  rule(s, 1.192, 3.843, 3.118);
}

function slide11(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 13.333, 2.676);
  header(s, TRUCK);
  text(s, "Delizioso's journey began with a passion for Mexican cuisine and a dream of sharing our " +
    'love for gourmet tacos and burritos with the world. Since our humble beginnings, we have ' +
    'PLACEHOLDER' +
    'innovative flavors. Our success is a testament to the hard work and commitment of our team, ' +
    'as well as the support of our valued customers.',
    lead({ x: 1.192, y: 5.757, w: 11.085, h: 0.825 }));
  text(s, 'Our Success Story', { x: 2.355, y: 4.573, w: 5.107, h: 0.774, fontSize: 40 });
  pill(s, 7.454, 3.532, 2.787, 0.522, SITE, { x: 8.038, y: 3.642, w: 1.618, size: 12 });
  arrowDisc(s, 11.912, 1.474, 0.492);
  rule(s, 9.154, 1.72, 2.219);
  logo(s, 1.192, 4.467, 0.984);
  text(s, 'Delizioso',
    { x: 11.01, y: 3.575, w: 1.413, h: 0.438, fontSize: 20, align: 'right', valign: 'middle' });
}

function slide12(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 6.241, 7.5);
  const drivers = [
    { y: 1.78, oy: 1.784, nx: 0.785, nw: 0.304,
      copy: 'Growing Demand for Unique and High-Quality Dining: The culinary landscape is evolving, ' +
        'with more customers seeking unique and elevated dining experiences. ' },
    { y: 3.478, oy: 3.483, nx: 0.757, nw: 0.353,
      copy: 'Increasing Popularity of Food Trucks: Food trucks have become a popular choice for ' +
        'quick and convenient meals, offering a vibrant and casual dining experience. ' },
    { y: 5.177, oy: 5.182, nx: 0.759, nw: 0.356,
      copy: 'Rising Interest in Mexican Cuisine: Mexican cuisine continues to gain popularity ' +
        'worldwide, with its vibrant flavors and rich cultural heritage. ' },
  ];
  drivers.forEach((d, i) => {
    text(s, d.copy, body({ x: 1.592, y: d.y, w: 3.129, h: 1.077 }));
    s.addShape('ellipse', {
      x: 0.67, y: d.oy, w: 0.534, h: 0.534, fill: { color: BLACK }, line: { type: 'none' },
    });
    text(s, String(i + 1), {
      x: d.nx, y: d.oy + 0.065, w: d.nw, h: 0.404,
      fontSize: 18, bold: true, color: WHITE, align: 'center', valign: 'middle',
    });
  });
  text(s, 'Market Opportunity', { x: 7.214, y: 1.375, w: 5.449, h: 0.774, fontSize: 40 });
  pill(s, 7.076, 3.032, 2.787, 0.522, SITE, { x: 7.66, y: 3.142, w: 1.618, size: 12 });
  text(s, 'PLACEHOLDER' +
    'High-quality Dining Experiences',
    note({ x: 8.738, y: 4.228, w: 3.934, h: 0.97 }));
  header(s, HEADLINE);
  arrowDisc(s, 12.201, 6.063, 0.492);
  rule(s, 8.881, 6.309, 2.563);
}

function slide13(pptx) {
  const s = pptx.addSlide();
  panel(s, 6.667, 1.191, 6.005, 5.118);
  text(s, 'Our Food Truck', { x: 9.181, y: 2.526, w: 1.457, h: 0.303 });
  text(s, 'We serve delicious tacos and burritos from our food truck, using fresh ingredients, ' +
    'providing unique dining experiences, and exceptional customer service. We want to grow by ' +
    'adding more trucks in different cities, catering private events, and opening a restaurant.',
    body({ x: 9.181, y: 2.887, w: 2.559, h: 2.087 }));
  text(s, 'Business Model',
    { x: 0.863, y: 3.327, w: 2.904, h: 1.582, fontSize: 44, align: 'right', wrap: true });
  header(s, null);
  arrowDisc(s, 3.258, 2.591, 0.492);
  rule(s, 0.639, 2.837, 2.14);
}

function slide14(pptx) {
  const s = pptx.addSlide();
  panel(s, 5.037, 5.325, 1.192, 2.175);
  panel(s, 0, 0, 13.333, 1.191);
  text(s, 'We have big goals to grow and expand. Our plans for the near future',
    lead({ x: -0.51, y: 4.581, w: 2.883, h: 0.572, rotate: 270 }));
  text(s, 'Plans to Grow', { x: 6.608, y: 3.493, w: 2.709, h: 1.582, fontSize: 44, wrap: true });
  const plans = [
    { y: 3.045, label: 'More Food Trucks: ',
      copy: "We'll have more food trucks in various cities to reach more customers." },
    { y: 4.271, label: 'Catering Services: ',
      copy: "We'll cater to private events, weddings, and corporate functions." },
    { y: 5.498, label: 'Permanent Location: ',
      copy: "We'll eventually have a fixed location in addition to our mobile food truck." },
  ];
  plans.forEach((p) => {
    text(s, [
      { text: p.label, options: { bold: true, color: BLACK } },
      { text: p.copy, options: { color: GRAY } },
    ], { x: 10.112, y: p.y, w: 2.46, h: 0.825, fontSize: 10, wrap: true, lineSpacingMultiple: LS });
  });
  header(s, HEADLINE);
  arrowDisc(s, 10.112, 2.089, 0.492);
  rule(s, 6.334, 2.335, 2.991);
}

function slide15(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 4.206, 7.5);
  text(s, [
    { text: 'Meet Our Founder ', options: { breakLine: true } },
    { text: '& Head Chef' },
  ], { x: 5.261, y: 5.653, w: 4.621, h: 1.313, fontSize: 36 });
  pill(s, 4.6, 4.267, 3.217, 0.693, 'Juan Rodriguez', { x: 5.312, y: 4.429, w: 1.792, size: 16 });
  text(s, 'With Years Of Experience In The Restaurant Industry, Juan Ensures That Every Dish Meets ' +
    'The Highest Standards Of Taste And Quality.',
    note({ x: 8.542, y: 2.397, w: 4.13, h: 0.97 }));
  text(s, "Juan's passion for Mexican cuisine and his extensive culinary expertise have been " +
    "instrumental in shaping Delizioso's menu and flavor profiles.",
    lead({ x: 7.288, y: 0.905, w: 5.334, h: 0.572 }));
  text(s, HEADLINE, { x: 0.652, y: 6.394, w: 2.38, h: 0.505, wrap: true });
  header(s, null);
  arrowDisc(s, 7.293, 2.71, 0.492);
  rule(s, 4.6, 2.956, 1.936);
}

function slide16(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 1.191, 1.4, 6.309);
  const team = [
    { y: 1.465, name: 'Maria Hernandez - Operations Manager',
      copy: 'Maria oversees the day-to-day operations of Delizioso, ensuring smooth logistics, ' +
        'efficient service, and exceptional customer experiences.' },
    { y: 3.211, name: 'Carlos Lopez - Customer Relations Specialist',
      copy: 'Carlos is the friendly face behind Delizioso, interacting with customers, addressing ' +
        'their inquiries, and building strong relationships.' },
    { y: 4.957, name: 'Juan Rodriguez - Founder and Head Chef',
      copy: "Juan'PLACEHOLDER's menu to " +
        'exceptional standards of taste and quality.' },
  ];
  team.forEach((m) => {
    text(s, [
      { text: m.name, options: { bold: true, color: BLACK, breakLine: true } },
      { text: m.copy, options: { color: GRAY } },
    ], { x: 8.236, y: m.y, w: 3.888, h: 1.077, fontSize: 10, wrap: true, lineSpacingMultiple: LS });
  });
  text(s, [
    { text: 'Meet Our Expert', options: { breakLine: true } },
    { text: 'Team' },
  ], { x: 2.458, y: 1.465, w: 4.083, h: 1.313, fontSize: 36 });
  header(s, TRUCK);
  arrowDisc(s, 6.541, 3.061, 0.492);
  rule(s, 2.483, 3.307, 3.396);
}

function slide17(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 1.843, 7.5);
  const people = [
    { nx: 3.461, ny: 1.593, nw: 1.461, first: 'Juan Rodriguez ', role: 'Head Chef',
      bx: 3.487, by: 2.305, bw: 3.075,
      copy: "Juan'PLACEHOLDER's menu to " +
        'exceptional standards of taste and quality,' },
    { nx: 3.498, ny: 4.345, nw: 1.173, first: 'Maria Perez ', role: 'Sous Chef',
      bx: 3.461, by: 5.107, bw: 3.191,
      copy: 'Maria is a skilled Mexican cook. She helped make some of their most famous dishes by ' +
        'being creative and careful.' },
    { nx: 9.442, ny: 1.609, nw: 1.89, first: 'Maria Hernandez ', role: 'Operations Manager',
      bx: 9.459, by: 2.289, bw: 3.075,
      copy: 'Maria oversees the day-to-day operations ensuring smooth logistics, efficient service, ' +
        'and exceptional customer experiences.' },
    { nx: 9.481, ny: 4.373, nw: 2.651, first: 'Carlos Lopez', role: 'Customer Relations Specialist',
      bx: 9.481, by: 5.08, bw: 3.191,
      copy: 'Carlos is the friendly face behind Delizioso, interacting with customers, addressing ' +
        'their inquiries, and building strong relationships.' },
  ];
  people.forEach((p) => {
    text(s, [
      { text: p.first, options: { breakLine: true } },
      { text: p.role },
    ], { x: p.nx, y: p.ny, w: p.nw, h: 0.505 });
    text(s, p.copy, lead({ x: p.bx, y: p.by, w: p.bw, h: 0.825 }));
  });
  header(s, TRUCK);
}

function slide18(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 0, 13.333, 2.946);
  const quotes = [
    { nx: 2.569, nw: 0.852, name: 'Laura M.', bx: 2.483, sx: 3.484,
      copy: '"Delizioso\'s gourmet tacos are a game-changer! The flavors are incredible, and the ' +
        'quality is unmatched. Every bite feels like a celebration of Mexican cuisine." ' },
    { nx: 8.607, nw: 0.845, name: 'David R.', bx: 8.607, sx: 9.521,
      copy: '"I can\'t get enough of Delizioso\'s burritos. The unique combinations and fresh ' +
        'ingredients make them a standout. It\'s like a flavor explosion in every bite!" ' },
  ];
  quotes.forEach((q) => {
    text(s, q.name, { x: q.nx, y: 4.223, w: q.nw, h: 0.303, align: 'right' });
    text(s, q.copy, lead({ x: q.bx, y: 4.691, w: 3.122, h: 1.077, italic: true }));
    for (let i = 0; i < 5; i++) {
      s.addShape('star5', {
        x: q.sx + i * 0.1972, y: 4.298, w: 0.153, h: 0.153,
        fill: { color: YELLOW }, line: { type: 'none' },
      });
    }
  });
  text(s, 'Testimonials', { x: 1.098, y: 1.738, w: 3.528, h: 0.774, fontSize: 40 });
  header(s, TRUCK);
  arrowDisc(s, 12.18, 6.653, 0.492);
  rule(s, 0.652, 6.899, 10.611);
}

function slide19(pptx) {
  const s = pptx.addSlide();
  panel(s, 0, 5.062, 13.333, 1.585);
  text(s, 'Get In Touch', { x: 6.125, y: 2.686, w: 4.176, h: 0.909, fontSize: 48 });
  text(s, 'For Inquiries, Bookings, Or to Learn More About Delizioso, Please Contact Us.',
    note({ x: 6.125, y: 3.887, w: 4.176, h: 0.667 }));
  const contacts = [
    { x: 6.125, lw: 0.724, vw: 1.194, label: 'Phone', value: '(555) 123-4567' },
    { x: 7.767, lw: 0.64, vw: 1.853, label: 'Email', value: ' info@Delizioso.website' },
    { x: 10.069, lw: 0.893, vw: 1.384, label: 'Website', value: SITE },
  ];
  contacts.forEach((c) => {
    text(s, c.label, { x: c.x, y: 5.556, w: c.lw, h: 0.303 });
    text(s, c.value, { x: c.x, y: 5.901, w: c.vw, h: 0.269, fontSize: 10, italic: true });
  });
  header(s, TRUCK);
  arrowDisc(s, 9.675, 1.439, 0.492);
  rule(s, 0.661, 1.685, 8.11);
}

function slide20(pptx) {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  panel(s, 0, 0, 2.778, 7.5);
  text(s, 'Thank You', { x: 6.859, y: 3.003, w: 5.221, h: 1.313, fontSize: 72, align: 'center' });
  logo(s, 5.544, 3.167, 0.984);
  text(s, "Thanks for checking out Delizioso! We'll serve you amazing tacos and burritos that will " +
    'keep you coming back for more. ',
    lead({ x: 5.544, y: 4.878, w: 2.87, h: 0.825 }));
  pill(s, 9.187, 4.878, 2.787, 0.522, SITE, { x: 9.771, y: 4.988, w: 1.618, size: 12 });
  header(s, HEADLINE);
  arrowDisc(s, 11.564, 2.19, 0.492);
  rule(s, 5.544, 2.437, 5.036);
}

// ---------------------------------------------------------------- assembly
function build() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE16x9', width: 13.333, height: 7.5 });
  pptx.layout = 'WIDE16x9';
  pptx.title = 'Delizioso Pitch Deck';

  [slide01, slide02, slide03, slide04, slide05, slide06, slide07, slide08, slide09, slide10,
    slide11, slide12, slide13, slide14, slide15, slide16, slide17, slide18, slide19, slide20]
    .forEach((fn) => fn(pptx));

  return pptx;
}

build().writeFile({
  fileName: path.join(__dirname, '17b26c2b-a18c-4ec9-ab35-f456ad1e2688_grok_final.pptx'),
}).then((f) => console.log('wrote', f));
