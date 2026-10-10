/**
 * Grok Bot page tokens, pinned from the Sand design prototype's
 * Figma-generated `globals.css` (light `:root` and dark `[data-theme]`
 * blocks) and its `.t-*` / `.w-*` type classes. Alpha-hex entries are ink at
 * the design's percent so they composite over any surface. `text.warning`
 * is the exception and comes from the Sand hue ramps instead: every yellow
 * step that holds 4.5:1 on the light page reads brown, so light takes the
 * most orange step that still passes (orange/light/10, Sand's
 * `text/supplementary1`), and dark takes Sand's `text/warning`
 * (yellow/dark/10). When the design tokens change, re-derive these; they do
 * not update automatically.
 */
/** One text style from the page type scale (Figma text styles on the system font). */
export interface GrokTextStyle {
    readonly fontSize: string;
    readonly lineHeight: string;
    readonly letterSpacing: string;
}
/** Ink and surface colors for one polarity. */
export interface GrokPalette {
    readonly bg: {
        /** Page background. */
        readonly base: string;
        /** Soft card behind the calendar. */
        readonly subtle: string;
    };
    readonly text: {
        readonly primary: string;
        /** Ink at 60%: summaries, table label column, details. */
        readonly secondary: string;
        /** Ink at 40%: captions, markers, axes, the rest of a chart. */
        readonly tertiary: string;
        /** Text on a checked control. */
        readonly onColor: string;
        readonly accent: string;
        readonly success: string;
        /** Only a genuinely bad status word; neutral states stay in neutral ink. */
        readonly warning: string;
        readonly danger: string;
    };
    readonly fill: {
        /** Solid ink fill. */
        readonly primary: string;
        /** Neutral 6% (light) / 8% (dark) tint: callouts, chart tracks. */
        readonly secondary: string;
        readonly accent: string;
        readonly controlChecked: string;
    };
    readonly border: {
        /** 5%: hairline rows. */
        readonly subtle: string;
        /** 10%: calendar hour lines. */
        readonly weak: string;
        /** 15%: link cards. */
        readonly default: string;
        /** 30%: unchecked checkbox. */
        readonly strong: string;
    };
}
/**
 * Grok Bot brand colors for data marks only: bars, columns, lines, points,
 * areas, slices, funnel bands, and legend swatches. Text never takes these;
 * labels, values, ticks, and legend names stay on the neutral `text` tokens.
 *
 * Hues are the Grok Bot brand midtones (step 9 of the Sand hue ramps). Dark
 * uses them as-is; light swaps in step 10 wherever step 9 is under 3:1
 * against the light page (WCAG 1.4.11 non-text contrast), so every
 * categorical color holds at least 3:1 on its own `bg.base`.
 *
 * Mirrored in `sand/src/shared/tokens/dataviz.ts` for the document renderer;
 * the two must stay in sync.
 */
export interface GrokDataViz {
    /**
     * Series colors in order: blue, orange, violet, green, magenta, yellow,
     * cyan, brown. Neighbors alternate warm and cool. Series 0 is the default
     * data color. Past eight series, fold the tail into `other`; never cycle.
     */
    readonly categorical: readonly [string, string, string, string, string, string, string, string];
    /** Gray for "Other", the rest, comparison baselines, and de-emphasized marks. */
    readonly other: string;
    /** Red, reserved for negative values and danger; not in the rotation. */
    readonly negative: string;
    /** The non-emphasized marks of a single series: a light tint of `categorical[0]`. */
    readonly singleSeriesRest: string;
    /** Alpha hex suffix (14%) for an area under a line: `${color}${areaAlpha}`. */
    readonly areaAlpha: "24";
    /** Alpha hex suffix (25%) for a funnel's previous-step band: `${color}${bandAlpha}`. */
    readonly bandAlpha: "40";
    /** Blue ramp, low to high, for ordered values (heatmaps, choropleths). */
    readonly sequential: readonly string[];
    /** Red, neutral, blue: below and above a midpoint. */
    readonly diverging: readonly string[];
}
