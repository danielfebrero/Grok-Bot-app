import { type JSX } from "react";
import { type DayEvent } from "./calendar-stack.js";
export type { DayEvent } from "./calendar-stack.js";
export interface DayViewProps {
    /** First hour shown, e.g. 9 or 8.5 for 8:30. */
    start: number;
    /** Last hour shown, e.g. 14, at most 24. Both hours are held to the day (0 to 24); events are clipped to `start..end`, and nothing renders when `end` is not after `start`. */
    end: number;
    events: readonly DayEvent[];
    /** One red line across the day, e.g. the current moment or a deadline, with a one-line label in a column at the right; skipped when `at` is outside `start..end`. The line breaks over an event's title row rather than striking through it. On a narrow block the label sits on the line above the events instead, the events on its row stop short of it, and it moves off any title it would cover. */
    mark?: {
        at: number;
        label: string;
    };
}
/**
 * A calendar day: a bordered card on the page, hour labels centered on their
 * lines, events as blue-tinted blocks with a 3 px blue rule down the left.
 * The selected event is solid blue with its title on the fill; free time is
 * a dashed outline; one danger-red line marks the moment, its label in a
 * column at the right that the events stop short of. Overlapping events
 * share the width in lanes. When the block is too narrow for lanes, events
 * stack under one another, folding what will not fit into a "+N more" row,
 * and the label sits on the line over them, or on the nearest clear row
 * when a title or the selected event's fill is in the way; an event on its
 * row stops short of it, as every event does on a wide block, so the label
 * sits beside the events rather than on their fills.
 */
export declare function DayView({ start: givenStart, end: givenEnd, events, mark, }: DayViewProps): JSX.Element | null;
