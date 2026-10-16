import { type JSX } from "react";
export interface LegendProps {
    /** Series names in order; name `i` takes `dataviz.categorical[i]`, and past three the gray `other`. */
    names: readonly string[];
    /** Indexes drawn as a dashed gray comparison (a baseline, last period, a target). */
    comparison?: readonly number[];
    /** Round swatches, for parts of a whole or a custom visual of points. */
    dots?: boolean;
}
/**
 * Names for the series of a chart: a colored swatch beside neutral text,
 * wrapping as a row above the visual it names, at the start of the line.
 * `ColumnSeries`, `LineChart`, and `Pie` draw their own, so never add one
 * to them; use it above a custom visual whose series take
 * `dataviz.categorical` in order, or above charts that set `legend={false}`
 * to share one.
 *
 * @kitNote Only above a custom visual; `ColumnSeries`, `LineChart`, and `Pie` draw their own.
 */
export declare function Legend({ names, comparison, dots }: LegendProps): JSX.Element;
export interface ReferenceLine {
    /** Where across the value axis the line sits, e.g. a target or a limit. */
    value: number;
    /** Short name printed at the line's right end, in a gutter the marks leave for it. */
    label: string;
}
export interface ColumnSeriesEntry {
    name: string;
    /** One value per label. */
    values: readonly number[];
}
export interface ColumnSeriesProps {
    /** The categories along the bottom, e.g. months or teams. Past about six, labels thin out; `""` never prints. */
    labels: readonly string[];
    /** Two or three series over `labels`, grouped side by side per label, or stacked; more share the gray tail. */
    series: readonly ColumnSeriesEntry[];
    /**
     * One column per label, the series stacked in order, with the total in
     * the detail. A stack is built from the parts at or above zero: a negative
     * part has no height in it (it still prints in the detail). Grouped
     * columns draw a negative value below the baseline in the negative red.
     */
    stacked?: boolean;
    /**
     * Top of the value axis; defaults to the next round number above the data
     * (or the tallest stack). Data above it still raises the axis, so a spike
     * is never flattened into a plateau.
     */
    ceiling?: number;
    /** Unit after each tick and detail figure, e.g. "%". Ignored when `format` is set. */
    unit?: string;
    /** How a tick and a detail figure read, e.g. `(v) => "$" + v.toLocaleString("en-US")`. */
    format?: (value: number) => string;
    /** A dashed marker across the plot, e.g. a target or a budget. */
    reference?: ReferenceLine;
    /**
     * Whether the chart prints its own legend above the columns (the default).
     * Set `false` only when one `Legend` placed by the author names the series
     * for this and a neighboring chart; never add a `Legend` to a chart that
     * still draws its own.
     */
    legend?: boolean;
}
/**
 * Several series as vertical columns per label, grouped or stacked, with a
 * legend of swatches and a detail box for the label under the pointer. For
 * one series use `Columns`; for parts of a whole across rows, `StackedBars`.
 * The legend is built in, above the plot: do not add a `Legend` to it.
 *
 * @kitNote Draws its own legend; add no `Legend` to it (`legend={false}` turns it off).
 *
 * @example
 * ```tsx
 * <ColumnSeries
 *   labels={["Q1", "Q2", "Q3", "Q4"]}
 *   series={[
 *     { name: "New", values: [42, 51, 47, 63] },
 *     { name: "Returning", values: [88, 92, 97, 104] },
 *   ]}
 *   stacked
 * />
 * ```
 */
export declare function ColumnSeries({ labels, series, stacked, ceiling, unit, format, reference, legend, }: ColumnSeriesProps): JSX.Element | null;
