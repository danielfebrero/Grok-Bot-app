import { type JSX, type ReactNode } from "react";
import type { Tone } from "./components.js";
export interface GridProps {
    /** Most cells per row, two by default. Cells wrap to fewer when the page is narrow; four fit a `wide` or `full` page. */
    columns?: 2 | 3 | 4;
    /**
     * One outline with hairlines between the cells instead of a card per cell;
     * a last cell short of a row then spans the rest of it.
     */
    joined?: boolean;
    /** No outline at all: the cells sit side by side with the gap between them, for small multiples that are already framed or need no frame. */
    plain?: boolean;
    /** The cells: blocks, custom visuals, or whole `Section`s. */
    children?: ReactNode;
}
/**
 * Two or three blocks side by side: small multiples, a chart next to its
 * breakdown, a before and after, the tiles of a dashboard. Each cell is a
 * card with a hairline outline; `joined` draws one outline with dividers
 * instead. Put it inside a `Section` to share one heading, or directly in
 * the `Page` with a `Section` per cell. On a `wide` or `full` page the cells
 * get the room of a column each. With no cells it draws nothing.
 *
 * @example
 * ```tsx
 * <Grid>
 *   <Section heading="Revenue by region"><Bars bars={regions} /></Section>
 *   <Section heading="Share of orders"><Pie parts={channels} /></Section>
 * </Grid>
 * ```
 */
export declare function Grid({ columns, joined, plain, children, }: GridProps): JSX.Element | null;
/** An option whose shown label differs from the value kept in state. */
export interface SegmentedOption<T extends string = string> {
    value: T;
    label: string;
    /** A count after the label in quiet tabular figures, e.g. how many rows the filter keeps. */
    count?: number;
}
/**
 * `T` is the option value type, inferred from `options`, `value`, and
 * `onChange` together, so a selection kept in a `useState<"a" | "b">` and a
 * handler typed on that union type-check without a cast.
 */
export interface SegmentedProps<T extends string = string> {
    /**
     * Two to five short labels, or `{ value, label }` pairs when the label
     * shown differs from the value kept in state.
     */
    options: readonly (T | SegmentedOption<T>)[];
    /** The selected value. */
    value: T;
    onChange: (value: T) => void;
}
/**
 * A quiet text switch for the view, metric, period, or filter a section
 * shows; two options are an on/off, up to five a pick. Keep the selection
 * in `useState` and compute each view from the embedded data. In a
 * `Section`'s `actions` it sits on the heading line.
 *
 * @example
 * ```tsx
 * const [metric, setMetric] = useState("Revenue");
 * <Segmented options={["Revenue", "Orders"]} value={metric} onChange={setMetric} />
 * ```
 */
export declare function Segmented<T extends string = string>({ options, value, onChange, }: SegmentedProps<T>): JSX.Element;
export interface Metric {
    /** The figure as it should read, e.g. "1.9%" or "$1.2M". */
    value: string;
    label: string;
    /** Quiet note under the label, e.g. "+12% vs Q2". */
    detail?: string;
    /** Colors `detail` when it is good or bad news. */
    tone?: Tone;
}
export interface MetricsProps {
    /**
     * Two to four headline figures. More still lay out, in equal rows, but
     * that many numbers read better in `Rows` or a `Table`.
     */
    metrics: readonly Metric[];
}
/**
 * A strip of headline figures on the page, no cards. It is the template
 * answer, so use it only when these few numbers are the point of the page;
 * otherwise put them in `Rows` or a `Table` next to what explains them. The
 * figures share the width equally, so their values and labels line up in
 * columns; when they do not all fit a row, they take even rows (four and
 * four, not seven and one), and on a phone four figures sit two by two. A
 * long figure widens every cell rather than breaking mid-number.
 */
export declare function Metrics({ metrics }: MetricsProps): JSX.Element;
