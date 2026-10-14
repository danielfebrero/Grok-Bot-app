import { type JSX, type ReactNode } from "react";
export interface QuoteProps {
    /** The quoted words, as said, in body text. */
    children?: ReactNode;
    /** Who said it or where it is from, one line under the words. */
    cite?: string;
}
/**
 * Someone else's words: a quiet bar beside body text, with an optional `cite`
 * line under it naming who said it or where it is from. For a quotation from
 * a customer, a message, a review, or a source; the page's own point is a
 * `Callout`. One with nothing to say draws nothing.
 */
export declare function Quote({ children, cite }: QuoteProps): JSX.Element | null;
