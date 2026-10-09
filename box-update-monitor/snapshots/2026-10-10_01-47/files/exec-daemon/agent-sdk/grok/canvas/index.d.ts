/**
 * Public API for authoring `.canvas.tsx` files via `grok/canvas`.
 *
 * A canvas built with this kit is a Grok Bot page: a title, a short summary,
 * and sections of paragraphs, callouts, lists, tables, and charts. The blocks
 * are the default vocabulary, not the ceiling: `Grid` sets blocks side by
 * side, `Segmented` switches views, `LineChart` (`lines`) and `ColumnSeries`
 * carry several series with a legend and a hover detail, and any richer visual (scatter,
 * heatmap, timeline, flow) is inline SVG or divs colored from
 * `useGrokTheme()` (`dataviz` for data marks, neutral `text` for every label,
 * type scale). A page has one width, set on the `Page`: the reading column
 * for a write-up, `wide` or `full` for a dashboard, a wide table, a board, or
 * a diagram, and every block shares it. It always looks like Grok and
 * follows the viewer's light/dark mode. Wrap the canvas in one `Page`, and
 * never hardcode colors or fonts.
 */
/**
 * React authoring utilities re-exported so canvases need one public import:
 * the hooks and the event, node, and style types a page's handlers and
 * helpers are typed with. A name missing here fails the canvas type check,
 * so the list covers what authors reach for, not only what the blocks use.
 */
export type { ChangeEvent, CSSProperties, FocusEvent, FormEvent, KeyboardEvent, MouseEvent, PointerEvent, ReactNode, RefObject, SyntheticEvent, } from "react";
export { Fragment, useCallback, useEffect, useMemo, useReducer, useRef, useState } from "react";
/** Calendar. */
export type { DayEvent, DayViewProps } from "./calendar.js";
export { DayView } from "./calendar.js";
/** Charts. */
export type { Bar, BarsProps, ColumnPoint, ColumnsProps, LineAnnotation, LineChartProps, LinePoint, LineSeries, PiePart, PieProps, StackedBar, StackedBarsProps, } from "./charts.js";
export { Bars, Columns, LineChart, Pie, StackedBars } from "./charts.js";
/** Several series as columns, and the legend a custom visual can share. */
export type { ColumnSeriesEntry, ColumnSeriesProps, LegendProps, ReferenceLine, } from "./charts-series.js";
export { ColumnSeries, Legend } from "./charts-series.js";
/** Layout math for a flow or dependency diagram the page draws itself. */
export type { DAGLayoutEdge, DAGLayoutNode, DAGLayoutOptions, DAGLayoutRank, DAGLayoutResult, } from "./dag.js";
export { computeDAGLayout } from "./dag.js";
/** Page shell, text, lists, table, links. */
export type { CalloutProps, ChecklistItem, ChecklistProps, Link, LinkCardsProps, ListProps, PageProps, Row, RowsProps, SectionProps, TableCell, TableProps, TextProps, Tone, } from "./components.js";
export { BulletList, Callout, Checklist, LinkCards, NumberedList, Page, Rows, Section, Table, Text, } from "./components.js";
/** Layout, view switching, and headline figures. */
export type { GridProps, Metric, MetricsProps, SegmentedOption, SegmentedProps } from "./layout.js";
export { Grid, Metrics, Segmented } from "./layout.js";
/** A field and a pressable for a tool page. */
export type { ButtonProps, InputProps } from "./controls.js";
export { Button, Input } from "./controls.js";
/** Page width. */
export type { PageWidth } from "./width.js";
/** Tokens. */
export type { GrokDataViz, GrokPalette, GrokTextStyle, GrokTheme } from "./hooks.js";
export { useGrokTheme } from "./hooks.js";
//# sourceMappingURL=index.d.ts.map