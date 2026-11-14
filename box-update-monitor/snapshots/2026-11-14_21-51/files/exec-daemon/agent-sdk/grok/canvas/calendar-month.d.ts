import type { JSX } from "react";
export interface MonthEvent {
    /** The event's day as `YYYY-MM-DD`; one outside the month is left out. */
    date: string;
    title: string;
}
export interface MonthViewProps {
    /** The month shown, as `YYYY-MM`. Nothing renders for any other text. */
    month: string;
    events: readonly MonthEvent[];
    /** The day drawn as today, `YYYY-MM-DD`: its number on a filled circle. */
    today?: string;
}
/**
 * A calendar month: the day view's card, a header of weekday names, Monday
 * first, and a hairline cell per day with its number at the top left,
 * today's on a filled circle. Cells before the first and after the last
 * day stay blank. An event is a chip in the day view's tint, a short blue
 * rule then its title on one caption line; a day with more than three shows
 * two and a "+N more" chip, and rows grow as tall as their fullest day. On
 * a narrow block the chips keep their rule and drop their title, which
 * stays in the tooltip.
 */
export declare function MonthView(props: MonthViewProps): JSX.Element | null;
