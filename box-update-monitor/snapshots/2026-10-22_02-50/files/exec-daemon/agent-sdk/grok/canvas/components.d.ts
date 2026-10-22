import { type JSX, type ReactNode } from "react";
import { type PageWidth } from "./width.js";
/**
 * The one output style every Grok Bot page uses. A page is a title, a short
 * summary, and sections of text, callouts, lists, tables, charts, and custom
 * visuals. Neutral text, brand-colored data marks: text sits on the type
 * scale in neutral ink, color in text marks only status, and chart marks take
 * the Grok Bot data-viz palette (`useGrokTheme().dataviz`). Every chart sits
 * in its own card, which the kit draws; no other pills, tints, cards, logos,
 * or decorative buttons inside the page.
 */
/** Color for a status word. No other text in a page is colored. */
export type Tone = "success" | "warning" | "danger";
export interface PageProps {
    /** Page title, once. */
    title: string;
    /** One or two sentences under the title: the point of the page. */
    summary?: string;
    /**
     * @deprecated Renders nothing. Accepted so pages written when it drew a
     * "From …" line still compile; leave it out of new pages.
     */
    sources?: readonly string[];
    /**
     * The one width every block of the page shares. `column` (the default) is
     * the 700 px reading measure, for a report, memo, or write-up; `wide` grows
     * with the window up to 1280 px, for a dashboard, a wide table, a board, or
     * a diagram; `full` spans the page. On a phone every width is the one column.
     */
    width?: PageWidth;
    /** `Section`s, each 64 px below what precedes it; a block placed directly here keeps a block's spacing. */
    children?: ReactNode;
}
/**
 * Root of a Grok canvas: title, summary, then sections, all at the page's
 * one `width`: the 700 px reading column by default, `wide` (up to 1280 px)
 * or `full` for a page that is mostly visuals. Wrap the whole canvas in one
 * `Page`.
 *
 * @example
 * ```tsx
 * <Page title="Fourth shop: where to open" summary="Mill Hill, opening in spring.">
 *   <Section heading="Pickup orders by neighborhood">…</Section>
 * </Page>
 * ```
 */
export declare function Page({ title, summary, width, children }: PageProps): JSX.Element;
export interface SectionProps {
    /** Section heading. Also names any chart inside; charts have no title of their own. */
    heading?: string;
    /**
     * The section's anchor, the target of a `TableOfContents` entry or an
     * `href="#…"`. By default the heading's words, lower case and hyphenated
     * (`"What changed"` is `what-changed`), numbered past any earlier id on
     * the page; set it only to name the anchor yourself, and it is used as
     * written.
     */
    id?: string;
    /**
     * One quiet line, e.g. the unit, the period, or a note on the heading.
     * Under the section's chart or table when it has one: inside a chart's
     * card, 16 px under the plot; 12 px under a table or any other figure. It
     * follows the first figure among the blocks, and a chart with nothing to
     * draw still shows it. In a grid cell it sits right under the cell's plot,
     * so give the cells of a row captions together or not at all if their
     * plots must line up. With no figure among the blocks it is the heading's
     * subtitle, right under it.
     */
    caption?: string;
    /** A control at the end of the heading line: a `Segmented` that switches the view, a `Button`. */
    actions?: ReactNode;
    /** The heading folds the blocks away; needs a `heading`. */
    collapsible?: boolean;
    /** Start a collapsible section open. Closed by default. */
    defaultOpen?: boolean;
    /**
     * Its blocks: paragraphs, a callout, a list, a table, a chart, a `Grid`, a
     * custom visual, or a `Section` for a sub-section under a smaller heading.
     */
    children?: ReactNode;
}
/**
 * One section of a page: a heading, its blocks, and an optional caption
 * under the first figure among them, or under the heading when there is no
 * figure, spaced by what they are: paragraphs 16 px apart, a list 24 px from
 * the prose around it, a figure 32 px, a strip of metrics 40 px, two charts
 * 16 px apart. Directly in the `Page` the heading is the section heading;
 * inside another `Section` it is the smaller sub-section heading; in a grid
 * cell it is the card's title, and the cell's blocks start under it at the
 * same gap as in every other card. `actions` end the heading line and wrap
 * under it on a narrow page; a `collapsible` section folds its blocks behind
 * a chevron, and a caption under a figure folds with them while a subtitle
 * stays with the heading and presses with it.
 */
export declare function Section(props: SectionProps): JSX.Element;
export interface TextProps {
    /** One paragraph of plain sentences. */
    children?: ReactNode;
}
/** A paragraph of body text. Use one `Text` per paragraph. */
export declare function Text(props: TextProps): JSX.Element;
/**
 * What a `Callout` is about. `info` is a note or a fact to keep in mind,
 * `success` what shipped or went well, `warning` what to watch, `danger`
 * what is blocked or broken. A tone tints the box in its hue, leads with its
 * icon, and inks the lead and body; without one the box stays neutral.
 */
export type CalloutTone = "info" | Tone;
export interface CalloutProps {
    /** Bold first line: the takeaway in a few words. */
    lead?: string;
    /**
     * Tints the box, adds the tone's icon, and inks the text: `info` blue,
     * `success` green, `warning` yellow, `danger` red. Leave it off for a
     * neutral takeaway.
     */
    tone?: CalloutTone;
    /** A line or two of body text. */
    children?: ReactNode;
}
/**
 * The page's one takeaway: a bold lead and a line or two in a box. Neutral by
 * default; a `tone` tints the box, leads with its icon, and inks the text.
 * Use at most one or two per page. One with nothing to say draws nothing.
 */
export declare function Callout(props: CalloutProps): JSX.Element | null;
export interface ChecklistItem {
    text: string;
    /** The recorded status: ticked and struck through when true. */
    done?: boolean;
    /** A note, e.g. a due time: right of the title, or under it on a phone. */
    meta?: string;
}
export interface ChecklistProps {
    /** The items with their recorded status: `done` is ticked. */
    items: readonly ChecklistItem[];
}
/**
 * Items with a tick box each, showing each item's recorded status: a ticked
 * item is struck through. The list is read-only on the page; a Page is a
 * document only its author updates, so a status changes when the author
 * edits the page, through the conversation or a routine.
 *
 * @kitNote Shows recorded status (`done`), as of when the page was written; the viewer cannot tick an item. Say so in copy ("status as of…"), and update the page when a status changes.
 */
export declare function Checklist(props: ChecklistProps): JSX.Element;
export interface Row {
    /** Short fixed-width lead, e.g. a time or a rank. */
    lead?: string;
    title: string;
    detail?: string;
    /** Right-aligned note, e.g. an owner, a due time, a count, a value. */
    meta?: string;
    /** A status word after the title; `tone` colors it. */
    status?: {
        label: string;
        tone?: Tone;
    };
}
export interface RowsProps {
    rows: readonly Row[];
}
/**
 * The detailed list: hairline rows with a lead, a title, a status word, a
 * detail line, and a right-aligned meta. Two or three numbers belong here
 * (value as `meta`), not in a chart.
 */
export declare function Rows(props: RowsProps): JSX.Element;
/**
 * One table cell: its text, a status word with a `tone` of `"success"`,
 * `"warning"`, or `"danger"` that colors it, or any React node (a kit
 * element, a number) when a word is not enough.
 */
export type TableCell = string | {
    text: string;
    tone?: Tone;
} | ReactNode;
export interface TableProps {
    /** Header row. The first column is the label column; leave it `""` when it needs no header. */
    columns: readonly string[];
    /**
     * One array per row: label first, then values. A column of figures, dates,
     * or single words is right-aligned on one line; a column with a phrase or a
     * sentence in it reads from the left and wraps. A row longer than `columns`
     * adds columns with no header rather than losing its cells.
     */
    rows: readonly (readonly TableCell[])[];
    /** Index of the column to set in medium weight, e.g. the recommended option. */
    emphasis?: number;
    /**
     * Alignment per column index. `auto`, the default, right-aligns a column
     * of figures and left-aligns words; set `end` or `start` to override one
     * column, e.g. a status word that should sit with the figures.
     */
    align?: readonly ("auto" | "start" | "end")[];
}
/**
 * A bordered grid: one 0.5 px outline with a rule between every cell, the
 * header a step smaller than the cells in medium weight, figures
 * right-aligned in tabular numerals, words left-aligned and wrapping. One
 * whose words are wider than its page scrolls sideways, fading at the edge
 * with more beyond it; from five columns it scrolls as a unit with every
 * label on one line.
 */
export declare function Table(props: TableProps): JSX.Element;
