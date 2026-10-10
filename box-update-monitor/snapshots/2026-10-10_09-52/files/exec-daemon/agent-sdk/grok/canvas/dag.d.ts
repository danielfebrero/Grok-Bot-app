/**
 * Layout math for a flow, dependency, or org diagram: positions for the
 * nodes, anchor points for the edges, a bounding box per rank. Drawing is
 * the page's: boxes and lines in inline SVG colored from `useGrokTheme()`
 * (`fill.secondary` or a hairline for a node, `border.strong` for an edge,
 * `dataviz` only for a highlighted path). Cycles are tolerated; a back edge
 * is flagged so it can be drawn dashed.
 *
 * The types are declared here because the kit's declarations are staged on
 * their own and may only reference siblings.
 */
export interface DAGLayoutOptions {
    /** Nodes to lay out. Only `id` is required. */
    nodes: Array<{
        id: string;
    }>;
    /** Directed edges. */
    edges: Array<{
        from: string;
        to: string;
    }>;
    /** Flow direction. Default `"vertical"` (top to bottom). */
    direction?: "vertical" | "horizontal";
    /** Node box width in px. Default 160. */
    nodeWidth?: number;
    /** Node box height in px. Default 40. */
    nodeHeight?: number;
    /** Gap between ranks (layers) in px. Default 64. */
    rankGap?: number;
    /** Gap between sibling nodes in the same rank in px. Default 48. */
    nodeGap?: number;
    /** Padding around the bounding box in px. Default 24. */
    padding?: number;
}
export interface DAGLayoutNode {
    id: string;
    /** Left edge of the node box. */
    x: number;
    /** Top edge of the node box. */
    y: number;
    /** Layer index (0 = root). */
    rank: number;
    /** Position within the rank (0-indexed). */
    order: number;
}
export interface DAGLayoutEdge {
    from: string;
    to: string;
    /** Suggested source anchor point (center of the outgoing side). */
    sourceX: number;
    sourceY: number;
    /** Suggested target anchor point (center of the incoming side). */
    targetX: number;
    targetY: number;
    /** True when this edge was identified as a back edge (part of a cycle). */
    isBackEdge: boolean;
}
export interface DAGLayoutRank {
    /** Rank index (0 = root). */
    rank: number;
    /** Left edge of the rank bounding box. */
    x: number;
    /** Top edge of the rank bounding box. */
    y: number;
    /** Width of the rank bounding box. */
    width: number;
    /** Height of the rank bounding box. */
    height: number;
    /** Node ids in this rank, in order. */
    nodeIds: string[];
}
export interface DAGLayoutResult {
    nodes: DAGLayoutNode[];
    edges: DAGLayoutEdge[];
    /** Bounding box per rank, for drawing layer bands. */
    ranks: DAGLayoutRank[];
    /** The direction used for this layout. */
    direction: "vertical" | "horizontal";
    /** Total width of the bounding box. */
    width: number;
    /** Total height of the bounding box. */
    height: number;
}
/**
 * Compute a hierarchical layout for a directed graph. Returns node
 * positions, edge anchor points, rank bounding boxes, and back-edge flags;
 * the page draws them.
 *
 * @example
 * ```tsx
 * const layout = computeDAGLayout({
 *   nodes: [{ id: "a" }, { id: "b" }, { id: "c" }],
 *   edges: [{ from: "a", to: "b" }, { from: "b", to: "c" }],
 * });
 * <svg width="100%" viewBox={`0 0 ${layout.width} ${layout.height}`}>
 *   {layout.edges.map((e) => <line key={e.from + e.to} x1={e.sourceX} y1={e.sourceY} x2={e.targetX} y2={e.targetY} stroke={theme.border.strong} strokeDasharray={e.isBackEdge ? "4 3" : undefined} />)}
 *   {layout.nodes.map((n) => <rect key={n.id} x={n.x} y={n.y} width={160} height={40} rx={8} fill={theme.fill.secondary} />)}
 * </svg>
 * ```
 */
export declare function computeDAGLayout(options: DAGLayoutOptions): DAGLayoutResult;
