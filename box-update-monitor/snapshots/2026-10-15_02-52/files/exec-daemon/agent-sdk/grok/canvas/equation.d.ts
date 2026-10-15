import { type JSX, type ReactNode } from "react";
export interface EquationProps {
    /**
     * The equation's number at the end of its line, `"(1)"`, only when the
     * text refers back to it by that number; most equations have none, and
     * none takes a word or a name.
     */
    label?: string;
    /**
     * The terms in reading order: strings for operands and operators (`"lift
     * ="`, `"≈ 16.2%"`, Unicode for symbols and simple scripts), and `Frac`,
     * `Pow`, `Sqrt`, `Sum`, or `Paren` for a layout a string cannot hold.
     */
    children?: ReactNode;
}
/**
 * One equation on its own centered line, its terms 10 px apart in body text,
 * with 16 px above and below. On a narrow column the terms wrap onto
 * centered lines. One with no terms draws nothing.
 *
 * @kitNote Display math only; a formula inside a sentence stays Unicode in `Text`. `label` is an equation number such as `"(1)"`, only when the text cites it; never a name. `Paren` around a fraction, root, or sum; plain `(` `)` in a string otherwise.
 * @example
 * ```tsx
 * <Equation label="(1)">
 *   lift = <Frac num="p₁ − p₀" den="p₀" /> ≈ 16.2%
 * </Equation>
 * ```
 */
export declare function Equation({ label, children }: EquationProps): JSX.Element | null;
export interface FracProps {
    /** Above the bar. */
    num: ReactNode;
    /** Below the bar. */
    den: ReactNode;
}
/** A fraction: numerator over a 1 px bar over denominator, centered on the bar, all at the surrounding size. */
export declare function Frac({ num, den }: FracProps): JSX.Element;
export interface PowProps {
    /** The term that carries the exponent. */
    base: ReactNode;
    /** The exponent, raised at 75% size. */
    exp: ReactNode;
}
/**
 * A base with its exponent raised beside it at 75% size: `(1 + r)` to the
 * `n`. Beside a one-line base the exponent sits in the top of the line; beside
 * a tall base (a fraction, a root, a sum, a `Paren` around one) its baseline
 * sits on the base's top edge, and the term's box grows by the lift above and
 * below, so the base stays centered on its line and the raised run never
 * leaves the box to meet the line above or the edge of a scrolling line.
 */
export declare function Pow({ base, exp }: PowProps): JSX.Element;
export interface ParenProps {
    /** The opening delimiter; `(` unless given. `(`, `)`, `[`, `]`, and `|` stretch; another character is set at body size. */
    open?: string;
    /** The closing delimiter; `)` unless given. */
    close?: string;
    /** The enclosed terms. */
    children?: ReactNode;
}
/**
 * Delimiters around a term, as tall as the term: around a fraction, a root,
 * or a sum each delimiter is a 1 px stroke in the text's ink that spans the
 * whole height; around one line of text each is the plain glyph, so
 * `<Paren>a + b</Paren>` reads exactly as the string `(a + b)`.
 */
export declare function Paren({ open, close, children }: ParenProps): JSX.Element;
export interface SqrtProps {
    /** The term under the root. */
    children?: ReactNode;
}
/**
 * A square root: the root sign, then its term under a 1 px rule. Around one
 * line of text the sign is the glyph; around a tall term (a fraction, a sum,
 * another root) it is a 1 px stroke as tall as the term that rises to meet
 * the rule.
 */
export declare function Sqrt({ children }: SqrtProps): JSX.Element;
export interface SumProps {
    /** The lower limit, under the sign at 75% size: `"t = 1"`. */
    from?: ReactNode;
    /** The upper limit, over the sign at 75% size: `"n"`. */
    to?: ReactNode;
    /** The sign itself; `∑` unless the sum is a product (`∏`) or an integral (`∫`). */
    sign?: string;
}
/** A sum sign with its limits stacked over and under it at 75% size; the summand follows as the next term. */
export declare function Sum({ from, to, sign }: SumProps): JSX.Element;
