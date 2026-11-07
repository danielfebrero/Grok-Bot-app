import type { JSX, ReactNode } from "react";
/** The color of a status pill: a tone for a state, `neutral` for one that is neither good nor bad. */
export type StatusTone = "success" | "warning" | "danger" | "accent" | "neutral";
export interface StatusProps {
    /** `success` (on track, done), `warning` (at risk), `danger` (blocked), `accent` (in progress), or `neutral`, the default. */
    tone?: StatusTone;
    /** The state in a word or two: `On track`, `At risk`, `Blocked`. */
    children: ReactNode;
}
/**
 * A state as a pill: a 6 px dot and a label in 12/16 medium, both in the
 * tone's ink, on the tone's tint, 20 px tall with 8 px of side padding and
 * fully rounded. It sits inline in a line of text, so a `Properties` value
 * or a table cell can carry one; it is not a block of its own. The pill
 * never grows past its line: a label too long for it ends in an ellipsis,
 * and the whole label is in the pill's tooltip.
 *
 * @example
 * ```tsx
 * <Properties items={[{ label: "Status", value: <Status tone="success">On track</Status> }]} />
 * ```
 */
export declare function Status({ tone, children }: StatusProps): JSX.Element;
