/**
 * The layout behind the calendar day and week views: the design's
 * measurements, the hour rows, and the placement of a column's events as
 * lanes or as a stack that folds what will not fit. Pure, no React.
 */
export interface DayEvent {
    /** Start hour, e.g. 9.5 for 9:30. */
    start: number;
    /** End hour. An event that ends before it starts is drawn as a point at its start. */
    end: number;
    title: string;
    /**
     * Kept for compatibility; it no longer changes the drawing. Every event
     * takes the day's one tinted style, so a flagged event looks like the rest.
     */
    flag?: boolean;
    /** Open time worth protecting, drawn as an outline. On a narrow block it gives way when events need the room. */
    free?: boolean;
    /** The event the page is about, drawn solid with its title on the fill. A held slot (`free`) stays an outline. */
    selected?: boolean;
}
export {};
