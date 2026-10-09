# grok/canvas kit reference (generated from the served declarations)
Import everything from `grok/canvas`. Components are listed as `<Name> { props }`; `?` marks an optional prop. Shared types follow the components that use them.
React re-exports: types ChangeEvent, CSSProperties, FocusEvent, FormEvent, KeyboardEvent, MouseEvent, PointerEvent, ReactNode, RefObject, SyntheticEvent; Fragment, useCallback, useEffect, useMemo, useReducer, useRef, useState.

## Page, text, lists, table, links
<BulletList> { items: string[] }
<Callout> { lead?: string; tone?: Tone; children?: ReactNode }
<Checklist> { items: ChecklistItem[]; value?: boolean[]; onChange?: (value: boolean[]) => void }
<LinkCards> { links: Link[]; numbered?: boolean }
<NumberedList> { items: string[] }
<Page> { title: string; summary?: string; sources?: string[]; width?: PageWidth; children?: ReactNode }
<Rows> { rows: Row[] }
<Section> { heading?: string; caption?: string; actions?: ReactNode; collapsible?: boolean; defaultOpen?: boolean; children?: ReactNode }
<Table> { columns: string[]; rows: (TableCell[])[]; emphasis?: number; align?: ("auto" | "start" | "end")[] }
<Text> { children?: ReactNode }
ChecklistItem { text: string; done?: boolean; meta?: string }
Link { title: string; site?: string; date?: string; href: string }
Row { lead?: string; title: string; detail?: string; meta?: string; status?: { label: string; tone?: Tone } }
TableCell = string | { text: string; tone?: Tone } | ReactNode
Tone = "success" | "warning" | "danger"
PageWidth = "column" | "wide" | "full"

## Layout, view switching, headline figures
<Grid> { columns?: 2 | 3 | 4; joined?: boolean; plain?: boolean; children?: ReactNode }
<Metrics> { metrics: Metric[] }
<Segmented<T extends string = string>> { options: (T | SegmentedOption<T>)[]; value: T; onChange: (value: T) => void }
Metric { value: string; label: string; detail?: string; tone?: Tone }
SegmentedOption<T extends string = string> { value: T; label: string; count?: number }

## Controls
<Button> { label: string; onClick: () => void }
<Input> { label?: string; value: string; onChange: (value: string) => void; placeholder?: string; type?: "text" | "number" }

## Charts
<Bars> { bars: Bar[]; emphasis?: number | "none" }
<Columns> { points: ColumnPoint[]; emphasis?: number | "none" }
<LineChart> { points: LinePoint[]; ceiling: number; floor?: number; unit?: string; format?: (value: number) => string; lines?: LineSeries[]; secondary?: { label: string; values: number[] }; labels?: [string, string?]; annotations?: LineAnnotation[] }
<Pie> { parts: PiePart[]; total?: boolean | number | string; label?: string }
<StackedBars> { series: string[]; bars: StackedBar[] }
Bar { label: string; value: number; display: string; meta?: string }
ColumnPoint { label: string; value: number; display?: string }
LineAnnotation { at: number | "zero"; label: string; line?: number; value?: number }
LinePoint { label: string; value: number }
LineSeries { label: string; values: number[] }
PiePart { label: string; value: number; display?: string }
StackedBar { label: string; values: number[]; display?: string }

## Series charts and legend
<ColumnSeries> { labels: string[]; series: ColumnSeriesEntry[]; stacked?: boolean; ceiling?: number; unit?: string; format?: (value: number) => string; reference?: ReferenceLine }
<Legend> { names: string[]; comparison?: number[]; dots?: boolean }
ColumnSeriesEntry { name: string; values: number[] }
ReferenceLine { value: number; label: string }

## Calendar
<DayView> { start: number; end: number; events: DayEvent[]; mark?: { at: number; label: string } }
DayEvent { start: number; end: number; title: string; flag?: boolean; free?: boolean }

## Flow / dependency layout
computeDAGLayout(options: DAGLayoutOptions): DAGLayoutResult
DAGLayoutEdge { from: string; to: string; sourceX: number; sourceY: number; targetX: number; targetY: number; isBackEdge: boolean }
DAGLayoutNode { id: string; x: number; y: number; rank: number; order: number }
DAGLayoutOptions { nodes: Array<{ id: string }>; edges: Array<{ from: string; to: string }>; direction?: "vertical" | "horizontal"; nodeWidth?: number; nodeHeight?: number; rankGap?: number; nodeGap?: number; padding?: number }
DAGLayoutRank { rank: number; x: number; y: number; width: number; height: number; nodeIds: string[] }
DAGLayoutResult { nodes: DAGLayoutNode[]; edges: DAGLayoutEdge[]; ranks: DAGLayoutRank[]; direction: "vertical" | "horizontal"; width: number; height: number }

## Theme
useGrokTheme(): GrokTheme
GrokDataViz { categorical: [string, string, string, string, string, string, string, string]; other: string; negative: string; singleSeriesRest: string; areaAlpha: "24"; bandAlpha: "40"; sequential: string[]; diverging: string[] }
GrokPalette { bg: { base: string; subtle: string }; text: { primary: string; secondary: string; tertiary: string; onColor: string; accent: string; success: string; warning: string; danger: string }; fill: { primary: string; secondary: string; accent: string; controlChecked: string }; border: { subtle: string; weak: string; default: string; strong: string } }
GrokTextStyle { fontSize: string; lineHeight: string; letterSpacing: string }
GrokTheme extends GrokPalette { kind: "light" | "dark"; font: { family: string; mono: string; heading1: GrokTextStyle; heading2: GrokTextStyle; heading3: GrokTextStyle; body1: GrokTextStyle; body2: GrokTextStyle; label: GrokTextStyle; caption: GrokTextStyle; code: GrokTextStyle }; weight: { regular: 400; medium: 500; semibold: 600 }; radius: { sm: 4; md: 8; lg: 12 }; dataviz: GrokDataViz; chart: [string, string, string, string] }
