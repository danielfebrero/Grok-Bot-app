import type { CSSProperties } from "react";
/**
 * A page has one width, chosen on the `Page` and shared by every block in it.
 * `column` is the reading measure for a report or memo; `wide` grows with the
 * window up to a cap, for a dashboard, a wide table, a board, or a diagram;
 * `full` spans the page. On a phone every width is the one column.
 */
export type PageWidth = "column" | "wide" | "full";
