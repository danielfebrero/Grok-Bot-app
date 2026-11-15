/**
 * Grok Bot page tokens, pinned from the Grok Bot PD System Figma variables
 * (`Semantic Colors` for ink and surfaces, `Data Colors` for marks, each in
 * its light and dark mode) and the page type styles. Alpha-hex entries are
 * ink at the design's percent so they composite over any surface. The values
 * are exact copies of the design's; when the design tokens change, re-derive
 * these, they do not update automatically.
 */
/** One text style from the page type scale (Figma text styles on the system font). */
export interface GrokTextStyle {
    readonly fontSize: string;
    readonly lineHeight: string;
    readonly letterSpacing: string;
    /**
     * A pinned variation-axis setting, where the design locks one: a face with
     * an `opsz` axis otherwise picks its optical size from the font size.
     */
    readonly fontVariationSettings?: string;
}
/** Ink and surface colors for one polarity. */
export interface GrokPalette {
    readonly bg: {
        /** Page background. */
        readonly base: string;
        /** A soft surface a step off the page: a held slot on the calendar. */
        readonly subtle: string;
        /** A surface that sits above the page: a chip in a chart, a popover. */
        readonly elevated: string;
    };
    readonly text: {
        readonly primary: string;
        /** Ink at 60%: summaries, table label column, details. */
        readonly secondary: string;
        /** Ink at 40%: captions, markers, axes, the rest of a chart. */
        readonly tertiary: string;
        /** Text on a checked control. */
        readonly onColor: string;
        /** Text on `fill.primary`: a filled button's label. */
        readonly onPrimary: string;
        readonly accent: string;
        readonly success: string;
        /** Only a genuinely bad status word; neutral states stay in neutral ink. */
        readonly warning: string;
        readonly danger: string;
        /**
         * Deep step of each tone's hue (`data/blue/5`, `data/green/5`,
         * `data/yellow/5`, `data/red/5`): body text inside a box tinted with the
         * matching `fill.*Subtle`, under a title in `accent` / `success` /
         * `warning` / `danger`. The ramp flips in dark so the step stays legible.
         */
        readonly accentDeep: string;
        readonly successDeep: string;
        readonly warningDeep: string;
        readonly dangerDeep: string;
    };
    readonly fill: {
        /** Solid ink fill. */
        readonly primary: string;
        /** `primary` under the pointer. */
        readonly primaryHover: string;
        /** Neutral 9% (light) / 17% (dark) tint: callouts, chart tracks, a quiet button under the pointer. */
        readonly secondary: string;
        readonly accent: string;
        readonly controlChecked: string;
        /**
         * Each tone's hue at the same 9% (light) / 17% (dark) as `secondary`
         * (`fill/accent-subtle`, `fill/success-subtle`, `fill/warning-subtle`,
         * `fill/danger-subtle`): the box of a toned callout or a tinted mark.
         */
        readonly accentSubtle: string;
        readonly successSubtle: string;
        readonly warningSubtle: string;
        readonly dangerSubtle: string;
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
        /** 40%: the ring on a focused control. */
        readonly focus: string;
    };
}
/**
 * Grok Bot data colors for data marks only: bars, columns, lines, points,
 * areas, slices, funnel bands, and legend swatches. Text never takes these;
 * labels, values, ticks, and legend names stay on the neutral `text` tokens.
 * The one exception the design draws is the title on a tinted calendar
 * event, which takes `blue5`.
 *
 * Three series hues and a gray tail, step 3 of the design's `Data Colors`
 * ramps, the same in light and dark. A single series uses the blue ramp: the
 * marks that matter at step 3, the rest at step 1, and the ramp flips in dark
 * so step 1 stays the quiet one. Red is never a series color; it means danger.
 *
 * The document renderer keeps its own copy of these values.
 */
export interface GrokDataViz {
    /**
     * Series colors in order: blue, yellow, cyan. Series 0 is the default data
     * color. Past three series, fold the tail into `other` and label it
     * directly; never cycle.
     */
    readonly categorical: readonly string[];
    /** Gray for "Other", the rest, comparison baselines, and de-emphasized marks. */
    readonly other: string;
    /** Red, reserved for negative values and danger; not in the rotation. */
    readonly negative: string;
    /** The non-emphasized marks of a single series: step 1 of the blue ramp. */
    readonly singleSeriesRest: string;
    /** Alpha hex suffix for an area under a line: `${color}${areaAlpha}`, the design's data tint (16% light, 24% dark). */
    readonly areaAlpha: string;
    /** Alpha hex suffix (25%) for a funnel's previous-step band: `${color}${bandAlpha}`. */
    readonly bandAlpha: string;
    /**
     * The design's `data/blue/tint`: step 3 of the blue ramp at the data tint
     * (16% light, 24% dark), the fill of a calendar event. Equal to
     * `${categorical[0]}${areaAlpha}`, pinned here under its own name.
     */
    readonly blueTint: string;
    /** The design's `data/blue/5`, the strong end of the blue ramp: the label on a tinted calendar event. */
    readonly blue5: string;
    /** Blue ramp, quiet to strong, for ordered values (heatmaps, choropleths). */
    readonly sequential: readonly string[];
    /** Red, neutral, blue: below and above a midpoint. */
    readonly diverging: readonly string[];
}
/** The page's weights; `regular` is 425 only for the brand variable face on a Mac. */
export type GrokWeightScale = {
    readonly regular: 400 | 425;
    readonly medium: 500;
    readonly semibold: 600;
};
