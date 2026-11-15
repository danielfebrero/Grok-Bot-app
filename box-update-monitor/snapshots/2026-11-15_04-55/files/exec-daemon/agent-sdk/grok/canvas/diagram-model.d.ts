/**
 * The model behind a `Diagram`: the node and edge props, the design's
 * measurements for node chrome and lines, and the measurement of a node
 * from its text. Pure, no React.
 */
/** The shape of a node. */
export type DiagramNodeKind = "box" | "decision" | "terminal" | "person";
export interface DiagramNode {
    /** Unique id, the one edges name in `from` and `to`. */
    id: string;
    /** The node's name; a person's name on a `person`. One line, cut with an ellipsis when long. */
    label: string;
    /** A second line under the label: a role on a `person`, a note on a `box`. A `box` with one draws as a card with an accent bar. */
    detail?: string;
    /** `box` (default): a card or a step. `decision`: a diamond. `terminal`: a pill for a start or end. `person`: an initials tile beside the name and role. */
    kind?: DiagramNodeKind;
    /** Initials in a `person`'s tile. Derived from the label when absent. */
    avatar?: string;
    /** Accent hue of a `box` card, as a series index into `dataviz.categorical` (0 blue, 1 yellow, 2 cyan); past three, and when absent, the gray `other`. */
    series?: number;
}
export interface DiagramEdge {
    /** Id of the node the line leaves. */
    from: string;
    /** Id of the node the line points to. */
    to: string;
    /** A short label beside the line: "Yes", "reports to", "owns". */
    label?: string;
    /** A second, fainter line under the label. */
    detail?: string;
    /** Draw the line dashed: a soft, optional, or planned relationship. */
    dashed?: boolean;
}
export interface DiagramProps {
    /** The nodes. Duplicate ids keep the first; an empty list draws nothing. */
    nodes: DiagramNode[];
    /** Directed edges between the nodes. One naming an unknown id, or a node itself, is skipped; two in opposite directions merge into one line with a head at each end, unless an end is a `decision`, where the way back draws as a loop; a cycle is tolerated. */
    edges: DiagramEdge[];
    /** Which way the ranks run. Default `"vertical"`, top to bottom; `"horizontal"` runs them left to right. */
    direction?: "horizontal" | "vertical";
}
