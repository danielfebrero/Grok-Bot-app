/**
 * Public API for authoring `.canvas.tsx` files via `grok/canvas`.
 *
 * A canvas built with this kit is a Grok Bot page: a title, a short summary,
 * and sections of paragraphs, callouts, lists, tables, and charts. The blocks
 * are the default vocabulary, not the ceiling: `Grid` sets blocks side by
 * side, `Segmented` switches views, `LineChart` (`lines`) and `ColumnSeries`
 * carry several series with a legend and a hover detail, `Scatter` puts two
 * measures per item in a cloud of dots, `Heatmap` and `CalendarHeatmap` color
 * a matrix or a year of days on the sequential ramp, `Diagram` draws a
 * flowchart or an org chart, and any richer visual (timeline) is inline SVG
 * or divs colored from
 * `useGrokTheme()` (`dataviz` for data marks, neutral `text` for every label,
 * type scale). A page has one width, set on the `Page`: the reading column
 * for a write-up, `wide` or `full` for a dashboard, a wide table, a board, or
 * a diagram, and every block shares it. It always looks like Grok and
 * follows the viewer's light/dark mode. A page is a document only its author
 * updates: a `Checklist` or `Checkbox` shows recorded status and takes no
 * input from the viewer, while `Segmented`, `Toggle`, `Select`, `Input`, and
 * `Button` are view controls the reader may operate to change what is shown,
 * never what is recorded. Wrap the canvas in one `Page`, and never hardcode
 * colors or fonts.
 */
/**
 * React authoring utilities re-exported so canvases need one public import:
 * the hooks and the event, node, and style types a page's handlers and
 * helpers are typed with. A name missing here fails the canvas type check,
 * so the list covers what authors reach for, not only what the blocks use.
 */
export type { ChangeEvent, CSSProperties, FocusEvent, FormEvent, KeyboardEvent, MouseEvent, PointerEvent, ReactNode, RefObject, SyntheticEvent, } from "react";
export { Fragment, useCallback, useEffect, useMemo, useReducer, useRef, useState } from "react";
/** People: one initials tile, and a stack of them with a `+N` overflow. */
export type { AvatarPerson, AvatarProps, AvatarSize, AvatarStackProps } from "./avatar.js";
export { Avatar, AvatarStack } from "./avatar.js";
/** Calendar: a day of events, one of them `selected`, with held slots and one marked moment; a week of them in a column per day, with today and `now`. */
export type { DayEvent, DayViewProps } from "./calendar.js";
export { DayView } from "./calendar.js";
export type { WeekEvent, WeekViewProps } from "./calendar-week.js";
export { WeekView } from "./calendar-week.js";
/** Calendar: a month of day cells, Monday first, with today and an event chip per event. */
export type { MonthEvent, MonthViewProps } from "./calendar-month.js";
export { MonthView } from "./calendar-month.js";
/** Charts. */
export type { Bar, BarsProps, ColumnPoint, ColumnsProps, LineAnnotation, LineChartProps, LinePoint, LineSeries, PiePart, PieProps, StackedBar, StackedBarsProps, } from "./charts.js";
export { Bars, Columns, LineChart, Pie, StackedBars } from "./charts.js";
/** Several series as columns, and the legend a custom visual can share. */
export type { ColumnSeriesEntry, ColumnSeriesProps, LegendProps, ReferenceLine, } from "./charts-series.js";
export { ColumnSeries, Legend } from "./charts-series.js";
/** Two measures per item as dots, sized by a third, with a dashed trend. */
export type { ScatterAxis, ScatterMeasure, ScatterPoint, ScatterProps } from "./scatter.js";
export { Scatter } from "./scatter.js";
/** Heatmaps: a matrix of rows by columns, and a calendar of one square per day. */
export type { CalendarDay, CalendarHeatmapProps, HeatmapProps } from "./heatmap.js";
export { CalendarHeatmap, Heatmap } from "./heatmap.js";
/** A 2×2 matrix and a Venn of two or three sets. */
export type { Matrix2x2Props, MatrixAxes, MatrixPoint, VennProps, VennSet, } from "./diagrams-simple.js";
export { Matrix2x2, Venn } from "./diagrams-simple.js";
/** Layout math for a flow or dependency diagram the page draws itself. */
export type { DAGLayoutEdge, DAGLayoutNode, DAGLayoutOptions, DAGLayoutRank, DAGLayoutResult, } from "./dag.js";
export { computeDAGLayout } from "./dag.js";
/** A drawn diagram on that layout: node cards, lines, arrowheads, labels, for a relationship diagram, a flowchart, an org chart. */
export type { DiagramEdge, DiagramNode, DiagramNodeKind, DiagramProps } from "./diagram.js";
export { Diagram } from "./diagram.js";
/** Page shell, text, callout, checklist, rows, table. */
export type { CalloutProps, CalloutTone, ChecklistItem, ChecklistProps, PageProps, Row, RowsProps, SectionProps, TableCell, TableProps, TextProps, Tone, } from "./components.js";
export { Callout, Checklist, Page, Rows, Section, Table, Text } from "./components.js";
/** Lists and steps. */
export type { BulletListProps, ListItem, ListMarker, ListProps, Step, StepsProps, } from "./lists.js";
export { BulletList, NumberedList, Steps } from "./lists.js";
/** Links that leave the page. */
export type { Link, LinkCardsProps } from "./link-cards.js";
export { LinkCards } from "./link-cards.js";
/** The facts about a thing as key–value rows. */
export type { PropertiesProps, PropertyItem } from "./properties.js";
export { Properties } from "./properties.js";
/** A pros and cons pair as two cards. */
export type { ProsConsProps } from "./pros-cons.js";
export { ProsCons } from "./pros-cons.js";
/** A state as an inline pill: a dot and a word on a tint. */
export type { StatusProps, StatusTone } from "./status.js";
export { Status } from "./status.js";
/** The page's outline: in-page links to its section headings. */
export type { TableOfContentsItem, TableOfContentsProps } from "./toc.js";
export { TableOfContents } from "./toc.js";
/** Someone else's words, and the rule between two parts of a page. */
export type { QuoteProps } from "./quote.js";
export { Quote } from "./quote.js";
export type { DividerProps } from "./divider.js";
export { Divider } from "./divider.js";
/** Layout, view switching, and headline figures. */
export type { GridColumns, GridProps, Metric, MetricsProps, SegmentedOption, SegmentedProps, } from "./layout.js";
export { Grid, Metrics, Segmented } from "./layout.js";
/** A field and a pressable for a tool page (the pressable is `quiet` (tinted) or `primary` (filled), one `primary` per row), the pick and switch of a filter row, and the tick, which shows a recorded status and is read-only on the page. */
export type { ButtonProps, ButtonVariant, CheckboxProps, InputProps, SelectOption, SelectProps, ToggleProps, } from "./controls.js";
export { Button, Checkbox, Input, Select, Toggle } from "./controls.js";
/** Display math: a centered equation line and the fraction, exponent, root, sum, and delimiters it is built from. */
export type { EquationProps, FracProps, ParenProps, PowProps, SqrtProps, SumProps, } from "./equation.js";
export { Equation, Frac, Paren, Pow, Sqrt, Sum } from "./equation.js";
/** Page width. */
export type { PageWidth } from "./width.js";
/** Tokens. */
export type { GrokDataViz, GrokPalette, GrokTextStyle, GrokTheme, GrokWeightScale, } from "./hooks.js";
export { useGrokTheme } from "./hooks.js";
//# sourceMappingURL=index.d.ts.map