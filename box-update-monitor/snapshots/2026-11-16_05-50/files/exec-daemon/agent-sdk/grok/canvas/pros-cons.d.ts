import type { JSX } from "react";
import { type ListItem } from "./lists.js";
export interface ProsConsProps {
    /** The upsides, one per item; each takes the green plus. */
    pros: readonly ListItem[];
    /** The downsides, one per item; each takes the red minus. */
    cons: readonly ListItem[];
    /** The first card's title, `Pros` by default. */
    prosHeading?: string;
    /** The second card's title, `Cons` by default. */
    consHeading?: string;
}
/**
 * A pros and cons pair: two cards side by side, `Pros` with a green plus
 * before each item and `Cons` with a red minus, each card as tall as its
 * own list. The cards are tighter than a `Grid` cell (a 12 px radius,
 * 12/16/16 of padding, the list 8 px under the title) and 16 px apart; on
 * the phone column they stack. A side with no items draws no card, and a
 * pair with neither draws nothing.
 *
 * @example
 * ```tsx
 * <ProsCons
 *   pros={["One page, fewer taps", "Easier to A/B test"]}
 *   cons={["Long page on small phones"]}
 * />
 * ```
 */
export declare function ProsCons(props: ProsConsProps): JSX.Element | null;
