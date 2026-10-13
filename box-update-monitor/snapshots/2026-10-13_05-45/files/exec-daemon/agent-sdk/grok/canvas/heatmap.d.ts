import { type JSX } from "react";
export interface HeatmapProps {
    /** Row labels, top to bottom, e.g. the days of the week. */
    rows: readonly string[];
    /** Column labels, left to right, e.g. the hours of the day. */
    columns: readonly string[];
    /** One array per row, one number per column: `values[r][c]`. A missing, NaN, zero or negative entry is the gray "nothing" cell. */
    values: readonly number[][];
    /** The value the strongest step stands for; by default the largest value in the data. */
    max?: number;
}
/**
 * A matrix heatmap in a chart card: row labels down the left, column labels
 * across the top, and one 27 px cell per row and column, colored on the
 * blue ramp against the largest value. A scale above the plot reads the
 * ramp from "Fewer" to "More". Cells share the block's width up to 54 px
 * each; on a block too narrow to give each one 12 px the matrix keeps that
 * width and scrolls sideways within its block, fading at the edge with more
 * beyond it. Labels never take a data color.
 *
 * @example
 * ```tsx
 * <Heatmap
 *   rows={["Mon", "Tue", "Wed"]}
 *   columns={["6a", "8a", "10a", "12p"]}
 *   values={[
 *     [3, 12, 18, 9],
 *     [4, 15, 21, 11],
 *     [0, 8, 14, 6],
 *   ]}
 * />
 * ```
 *
 * @kitNote `Heatmap` is a matrix; a row of day squares across a year is `CalendarHeatmap`.
 */
export declare function Heatmap(props: HeatmapProps): JSX.Element | null;
export interface CalendarDay {
    /** The day as `YYYY-MM-DD`; anything else is skipped. Two entries on one day add up. */
    date: string;
    /** NaN, zero or negative is the gray "nothing" square. */
    value: number;
}
export interface CalendarHeatmapProps {
    /** One entry per day with a value; days in the range without one are gray. */
    days: readonly CalendarDay[];
    /** First day drawn, `YYYY-MM-DD`; by default the earliest day in `days`. */
    from?: string;
    /** Last day drawn, `YYYY-MM-DD`; by default the latest day in `days`. */
    to?: string;
    /** The value the strongest step stands for; by default the largest value in the data. */
    max?: number;
    /** A line under the plot, e.g. "212 days with a shipped change", left of the scale. */
    note?: string;
}
/**
 * A calendar heatmap in a chart card: one 12 px square per day, a column
 * per week from Monday down to Sunday, colored on the blue ramp against the
 * largest value, gray where a day has nothing. Month labels sit over the
 * week each month starts in, "Mon", "Wed" and "Fri" down the left. Under the
 * plot, a note on the left and the scale from "Less" to "More" on the right.
 * On a block too narrow for every week the calendar wraps into strips of
 * the weeks that fit, one under another. Labels never take a data color.
 *
 * @example
 * ```tsx
 * <CalendarHeatmap
 *   days={[
 *     { date: "2024-01-02", value: 3 },
 *     { date: "2024-01-03", value: 1 },
 *   ]}
 *   from="2024-01-01"
 *   to="2024-12-31"
 *   note="212 days with a shipped change"
 * />
 * ```
 *
 * @kitNote Pass `from` and `to` to draw a whole year; without them the calendar spans the days given.
 */
export declare function CalendarHeatmap({ days, from, to, max, note, }: CalendarHeatmapProps): JSX.Element | null;
export {};
