# grok/canvas kit reference (generated from the served declarations)
Import everything from `grok/canvas`. Components are listed as `<Name> { props }`; `?` marks an optional prop. A note after `—` is a rule for using that component. Shared types follow the components that use them.
React re-exports: types ChangeEvent, CSSProperties, FocusEvent, FormEvent, KeyboardEvent, MouseEvent, PointerEvent, ReactNode, RefObject, SyntheticEvent; Fragment, useCallback, useEffect, useMemo, useReducer, useRef, useState.

## Page, text, lists, table, links
<Callout> { lead?: string; tone?: CalloutTone; children?: ReactNode }
<Checklist> { items: ChecklistItem[] } — Shows recorded status (`done`), as of when the page was written; the viewer cannot tick an item. Say so in copy ("status as of…"), and update the page when a status changes.
<Page> { title: string; summary?: string; sources?: string[]; width?: PageWidth; children?: ReactNode }
<Rows> { rows: Row[] }
<Section> { heading?: string; id?: string; caption?: string; actions?: ReactNode; collapsible?: boolean; defaultOpen?: boolean; children?: ReactNode }
<Table> { columns: string[]; rows: (TableCell[])[]; emphasis?: number; align?: ("auto" | "start" | "end")[] }
<Text> { children?: ReactNode }
<TableOfContents> { items: TableOfContentsItem[]; title?: string }
CalloutTone = "info" | Tone
ChecklistItem { text: string; done?: boolean; meta?: string }
Row { lead?: string; title: string; detail?: string; meta?: string; status?: { label: string; tone?: Tone } }
TableCell = string | { text: string; tone?: Tone } | ReactNode
Tone = "success" | "warning" | "danger"
TableOfContentsItem { label: string; id?: string; depth?: number }
PageWidth = "column" | "wide" | "full"

## Quote, divider, properties, pros and cons, status
<Quote> { children?: ReactNode; cite?: string }
<Divider> {}
<Properties> { items: PropertyItem[] }
<ProsCons> { pros: ListItem[]; cons: ListItem[]; prosHeading?: string; consHeading?: string }
<Status> { tone?: StatusTone; children: ReactNode }
PropertyItem { label: string; value: ReactNode }
StatusTone = "success" | "warning" | "danger" | "accent" | "neutral"

## People
<Avatar> { name: string; size?: AvatarSize; neutral?: boolean } — Alone on a card or in a cell; several people are an `AvatarStack`.
<AvatarStack> { people: AvatarPerson[]; max?: number; size?: AvatarSize; label?: string } — Also a `Table` cell; one person is an `Avatar`.
AvatarPerson { name: string }
AvatarSize = 24 | 28 | 32

## Layout, view switching, headline figures
<Grid> { columns?: GridColumns; joined?: boolean; plain?: boolean; children?: ReactNode }
<Metrics> { metrics: Metric[] }
<Segmented<T extends string = string>> { options: (T | SegmentedOption<T>)[]; value: T; onChange: (value: T) => void }
GridColumns = 2 | 3 | 4 | 5 | 6 | 7
Metric { value: string; label: string; detail?: string; tone?: Tone }
SegmentedOption<T extends string = string> { value: T; label: string; count?: number }

## Controls
<Button> { label: string; onClick: () => void; variant?: ButtonVariant; disabled?: boolean } — `variant="primary"` fills the one action a row leads with; the rest stay on the quiet default.
<Checkbox> { label?: string; value: boolean; disabled?: boolean } — Shows recorded status; the viewer cannot tick it. Phrase copy as status ("as of…"), and edit the page when it changes.
<Input> { label?: string; value: string; onChange: (value: string) => void; placeholder?: string; type?: "text" | "number" }
<Select<T extends string = string>> { label?: string; options: (T | SelectOption<T>)[]; value: T; onChange: (value: T) => void; placeholder?: string; disabled?: boolean }
<Toggle> { label?: string; value: boolean; onChange: (value: boolean) => void; disabled?: boolean }
ButtonVariant = "primary" | "quiet"
SelectOption<T extends string = string> { value: T; label: string }

## Charts
<Bars> { bars: Bar[]; emphasis?: number | "none" }
<Columns> { points: ColumnPoint[]; emphasis?: number | "none" }
<LineChart> { points: LinePoint[]; ceiling: number; floor?: number; unit?: string; format?: (value: number) => string; lines?: LineSeries[]; secondary?: { label: string; values: number[] }; labels?: [string, string?]; annotations?: LineAnnotation[]; legend?: boolean } — Draws its own legend from `labels` and `lines`; add no `Legend` to it (`legend={false}` turns it off).
<Pie> { parts: PiePart[]; total?: boolean | number | string; label?: string }
<StackedBars> { series: string[]; bars: StackedBar[] }
<CalendarHeatmap> { days: CalendarDay[]; from?: string; to?: string; max?: number; note?: string } — Pass `from` and `to` to draw a whole year; without them the calendar spans the days given.
<Heatmap> { rows: string[]; columns: string[]; values: number[][]; max?: number } — `Heatmap` is a matrix; a row of day squares across a year is `CalendarHeatmap`.
Bar { label: string; value: number; display: string; meta?: string }
ColumnPoint { label: string; value: number; display?: string }
LineAnnotation { at: number | "zero"; label: string; line?: number; value?: number }
LinePoint { label: string; value: number }
LineSeries { label: string; values: number[] }
PiePart { label: string; value: number; display?: string }
StackedBar { label: string; values: number[]; display?: string }
CalendarDay { date: string; value: number }

## Series charts and legend
<ColumnSeries> { labels: string[]; series: ColumnSeriesEntry[]; stacked?: boolean; ceiling?: number; unit?: string; format?: (value: number) => string; reference?: ReferenceLine; legend?: boolean } — Draws its own legend; add no `Legend` to it (`legend={false}` turns it off).
<Legend> { names: string[]; comparison?: number[]; dots?: boolean } — Only above a custom visual; `ColumnSeries`, `LineChart`, and `Pie` draw their own.
ColumnSeriesEntry { name: string; values: number[] }
ReferenceLine { value: number; label: string }

## Scatter
<Scatter> { points: ScatterPoint[]; x?: ScatterAxis; y?: ScatterAxis; size?: ScatterMeasure; trend?: boolean; legend?: boolean } — Draws its own legend from the dots' `series`; add no `Legend` to it (`legend={false}` turns it off).
ScatterAxis extends ScatterMeasure { floor?: number; ceiling?: number }
ScatterMeasure { name?: string; unit?: string; format?: (value: number) => string }
ScatterPoint { x: number; y: number; label?: string; size?: number; series?: string }

## Calendar
<DayView> { start: number; end: number; events: DayEvent[]; mark?: { at: number; label: string } }
<MonthView> { month: string; events: MonthEvent[]; today?: string }
DayEvent { start: number; end: number; title: string; flag?: boolean; free?: boolean; selected?: boolean }
MonthEvent { date: string; title: string }

## Diagram, matrix, Venn, and flow / dependency layout
<Diagram> { nodes: DiagramNode[]; edges: DiagramEdge[]; direction?: "horizontal" | "vertical" }
<Matrix2x2> { axes: MatrixAxes; quadrants?: [string, string, string, string]; points: MatrixPoint[] }
<Venn> { sets: VennSet[]; overlap?: number | string }
computeDAGLayout(options: DAGLayoutOptions): DAGLayoutResult
DiagramEdge { from: string; to: string; label?: string; detail?: string; dashed?: boolean }
DiagramNode { id: string; label: string; detail?: string; kind?: DiagramNodeKind; avatar?: string; series?: number }
DiagramNodeKind = "box" | "decision" | "terminal" | "person"
MatrixAxes { x: [string, string]; y?: [string, string] }
MatrixPoint { label: string; x: number; y: number }
VennSet { label: string; count?: number | string; items?: string[] }
DAGLayoutEdge { from: string; to: string; sourceX: number; sourceY: number; targetX: number; targetY: number; isBackEdge: boolean }
DAGLayoutNode { id: string; x: number; y: number; rank: number; order: number }
DAGLayoutOptions { nodes: Array<{ id: string }>; edges: Array<{ from: string; to: string }>; direction?: "vertical" | "horizontal"; nodeWidth?: number; nodeHeight?: number; rankGap?: number; nodeGap?: number; padding?: number }
DAGLayoutRank { rank: number; x: number; y: number; width: number; height: number; nodeIds: string[] }
DAGLayoutResult { nodes: DAGLayoutNode[]; edges: DAGLayoutEdge[]; ranks: DAGLayoutRank[]; direction: "vertical" | "horizontal"; width: number; height: number }

## Display math
<Equation> { label?: string; children?: ReactNode } — Display math only; a formula inside a sentence stays Unicode in `Text`. `label` is an equation number such as `"(1)"`, only when the text cites it; never a name. `Paren` around a fraction, root, or sum; plain `(` `)` in a string otherwise.
<Frac> { num: ReactNode; den: ReactNode }
<Paren> { open?: string; close?: string; children?: ReactNode }
<Pow> { base: ReactNode; exp: ReactNode }
<Sqrt> { children?: ReactNode }
<Sum> { from?: ReactNode; to?: ReactNode; sign?: string }

## Theme
useGrokTheme(): GrokTheme
GrokDataViz { categorical: string[]; other: string; negative: string; singleSeriesRest: string; areaAlpha: string; bandAlpha: string; blueTint: string; blue5: string; sequential: string[]; diverging: string[] }
GrokPalette { bg: { base: string; subtle: string; elevated: string }; text: { primary: string; secondary: string; tertiary: string; onColor: string; onPrimary: string; accent: string; success: string; warning: string; danger: string; accentDeep: string; successDeep: string; warningDeep: string; dangerDeep: string }; fill: { primary: string; primaryHover: string; secondary: string; accent: string; controlChecked: string; accentSubtle: string; successSubtle: string; warningSubtle: string; dangerSubtle: string }; border: { subtle: string; weak: string; default: string; strong: string; focus: string } }
GrokTextStyle { fontSize: string; lineHeight: string; letterSpacing: string; fontVariationSettings?: string }
GrokTheme extends GrokPalette { kind: "light" | "dark"; font: { family: string; mono: string; heading1: GrokTextStyle; heading2: GrokTextStyle; heading3: GrokTextStyle; body1: GrokTextStyle; subtext: GrokTextStyle; table: GrokTextStyle; metric: GrokTextStyle; body2: GrokTextStyle; label: GrokTextStyle; caption: GrokTextStyle; code: GrokTextStyle }; weight: GrokWeightScale; radius: { sm: 4; md: 8; lg: 12; xl: 16; full: 9999 }; dataviz: GrokDataViz; chart: [string, string, string, string] }
GrokWeightScale = { regular: 400 | 425; medium: 500; semibold: 600 }

## calendar-week
<WeekView> { week: string; start: number; end: number; events: WeekEvent[]; today?: string; now?: number }
WeekEvent extends DayEvent { date: string }

## link-cards
<LinkCards> { links: Link[]; numbered?: boolean }
Link { title: string; site?: string; date?: string; href: string }

## lists
<BulletList> { marker?: ListMarker }
<NumberedList> { items: ListItem[] }
<Steps> { steps: Step[] }
ListItem = string | { text: ReactNode; items?: ListItem[]; bullets?: boolean }
ListMarker = "pro" | "con"
Step { title: ReactNode; detail?: ReactNode }
