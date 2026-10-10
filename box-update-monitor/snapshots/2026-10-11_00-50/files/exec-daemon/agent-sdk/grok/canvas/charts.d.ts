import { type JSX } from "react";
/**
 * Charts for the common shapes: neutral text, brand-colored data marks. Marks
 * take `useGrokTheme().dataviz`; every label, value, and tick stays on the
 * neutral `text` tokens. A single series is `categorical[0]`, the mark that
 * matters at full strength and the rest in `singleSeriesRest`; parts of a
 * whole take the categorical hues in order, named by a legend of colored
 * swatches beside neutral text. No chart has a title or an axis title: the
 * `Section` heading names it. For shapes these don't cover, draw SVG with the
 * same `dataviz` tokens.
 */
export interface Bar {
    label: string;
    value: number;
    /** The value as it should read, e.g. "58" or "$1.2M". */
    display: string;
    /** Quiet note after the value, e.g. a share. */
    meta?: string;
}
export interface BarsProps {
    /** Comparable values, largest first; any count from two up reads well. */
    bars: readonly Bar[];
    /**
     * Index of the one bar at full strength; the rest take the series' light
     * tint. Defaults to the largest; `"none"` gives every bar full strength.
     */
    emphasis?: number | "none";
}
/**
 * Horizontal bars, one 6 px track per row. The label column fits its longest
 * label up to two fifths of the block, then truncates. A negative value
 * draws its size in the `negative` red; its `display` carries the sign. On a
 * narrow block, or when a value and its note would take more than two fifths
 * of the block on one line, the note takes its own row under the bar, at
 * most two lines; a value wider than its column is cut with an ellipsis; the
 * track never goes under 48 px (32 px on a narrow block). A reading that is not a finite number has no say in the
 * scale and shows an empty track with a dash. Nothing renders for an empty
 * series.
 */
export declare function Bars({ bars, emphasis }: BarsProps): JSX.Element | null;
export interface StackedBar {
    label: string;
    /** One value per entry in `series`, none below zero; extra values are dropped, a negative counts as zero. */
    values: readonly number[];
    /** The total as it should read past the bar, e.g. "62%"; defaults to the sum in figures. */
    display?: string;
}
export interface StackedBarsProps {
    /**
     * Names of the parts, in the order of each bar's `values`. Past eight, the
     * rest share gray; a nameless part draws but has no legend entry.
     */
    series: readonly string[];
    bars: readonly StackedBar[];
}
/**
 * Horizontal bars split into the same parts, with a legend. The longest bar
 * sets the scale; each total sits just past its bar's end. Only the parts the
 * legend names draw, each at least zero; a part that is not a finite number
 * counts as zero, and a row with none readable prints a dash. The track never
 * goes under 48 px. Nothing renders for an empty series.
 */
export declare function StackedBars({ series, bars }: StackedBarsProps): JSX.Element | null;
export interface ColumnPoint {
    label: string;
    value: number;
    /** The value as it should read above the column; defaults to the value in figures. */
    display?: string;
}
export interface ColumnsProps {
    /** Points of a short series, e.g. days of a week; any count. */
    points: readonly ColumnPoint[];
    /**
     * Index of the one column at full strength; the rest take the series' light
     * tint. Defaults to the largest; `"none"` gives every column full strength.
     */
    emphasis?: number | "none";
}
/**
 * Vertical bars in a 132 px grid, each at most 28 px wide, rising from a
 * zero baseline; a negative value hangs below it in the `negative` red.
 * A series denser than the block can draw at 2 px a column is binned to
 * runs of neighbors, each reading as its first label and largest reading.
 * When the block is too narrow for every value or label to print, they thin
 * to the ones that fit, the emphasized column's always among them, a label
 * wider than its room is cut with an ellipsis, and a dense series gives up
 * gap before column width. A reading that is not a
 * finite number has no say in the scale and shows no bar and a dash. Nothing
 * renders for an empty series.
 */
export declare function Columns({ points: given, emphasis }: ColumnsProps): JSX.Element | null;
export interface LinePoint {
    label: string;
    value: number;
}
/** Another line over the same points, in the next series color. */
export interface LineSeries {
    /** Its legend name. */
    label: string;
    /** One value per point. */
    values: readonly number[];
}
/** A short note pinned to one place on a line, offset from it with a leader. */
export interface LineAnnotation {
    /**
     * Where it sits: the index of a point (a fraction sits between two), or
     * `"zero"` for the first place the line crosses zero.
     */
    at: number | "zero";
    /** What it says, a few words: "Break even at 38 nights". */
    label: string;
    /** The line it marks: 0 is `points`, 1 the first of `lines`. Defaults to 0. */
    line?: number;
    /** A y value to pin it to instead of the line's value there. */
    value?: number;
}
export interface LineChartProps {
    points: readonly LinePoint[];
    /** Top of the y axis. Ticks at the floor, 0 when the floor is below it, half, and the ceiling. */
    ceiling: number;
    /**
     * Bottom of the y axis. Defaults to 0, or to a round number under the
     * lowest value when a line dips below zero, so a loss still plots. Never
     * hides a reading: a floor above the data drops to the lowest value.
     */
    floor?: number;
    /**
     * Unit after each tick, e.g. "ms". A long one prints once under the axis
     * instead. Ignored when `format` is set.
     */
    unit?: string;
    /** How a tick reads, e.g. `(v) => "$" + v.toLocaleString("en-US")`. */
    format?: (value: number) => string;
    /** More lines over the same points, in the series colors after the first. */
    lines?: readonly LineSeries[];
    /** A dashed gray comparison line, one value per point. */
    secondary?: {
        label: string;
        values: readonly number[];
    };
    /**
     * Names for the main line and the comparison, printed under the chart
     * beside their swatches. With `lines`, the legend prints whenever either
     * is set; a missing first name leaves the main line unnamed.
     */
    labels?: readonly [string, string?];
    /** Notes pinned to places on the lines, kept clear of the line and of each other. */
    annotations?: readonly LineAnnotation[];
}
/**
 * One line in the series color over a soft area of the same hue, with
 * optional further lines in the next series colors and a dashed gray
 * comparison. It draws in CSS pixels at the width of its box, never through a
 * scaled `viewBox`, so ticks and notes stay at label size on a wide page and
 * a phone alike. The y gutter widens to fit its longest tick label, and the
 * last x label ends at the chart's edge. It prints at most about six x-axis
 * labels, the first and last always; a point whose `label` is `""` never
 * prints one. A reading that is not a finite number is a gap in its line and
 * has no say in the axis; so has a `floor` or `ceiling` that is not one. A
 * long `unit` prints once under the axis instead of on every tick when the
 * gutter would otherwise take more than about a third of the chart. Nothing
 * renders for an empty series.
 */
export declare function LineChart({ points, ceiling, floor: floorProp, unit, format, lines, secondary, labels, annotations, }: LineChartProps): JSX.Element | null;
export interface PiePart {
    label: string;
    /** Its size, at least zero; a negative counts as zero. */
    value: number;
    /**
     * The value as it should read beside the label, e.g. "17%" or "540 orders";
     * defaults to the value in figures. When it is itself a percentage the
     * legend prints no share of its own after it.
     */
    display?: string;
}
export interface PieProps {
    /** Parts, largest first. Past eight, the rest share gray; fold small parts into one "Other". */
    parts: readonly PiePart[];
    /**
     * Print a total in the hole: `true` for the plain sum, a number to print
     * in figures, or the figure as it should read ("$1.2M"). A `label` under
     * it names it ("orders").
     */
    total?: boolean | number | string;
    /** A word under the total in the hole, e.g. "tickets". */
    label?: string;
}
/**
 * A donut of parts in the categorical hues, with each part's count and share
 * beside it (no share when the display already is one). A part whose value
 * is not a finite number is dropped. Pointing at a part or its legend row
 * lifts that part and quiets the rest; `total` prints the sum in the hole.
 * A long part label keeps to one line and truncates while the value and
 * share never break or shrink, whatever wrapping the page allows; on a
 * narrow block the legend moves under the donut so labels get the full
 * width first. Nothing renders for an empty series.
 */
export declare function Pie({ parts: given, total: showTotal, label }: PieProps): JSX.Element | null;
export {};
