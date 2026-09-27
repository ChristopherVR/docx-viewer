/**
 * Engine-side input contracts. These are intentionally a separate type family
 * from `@christophervr/docx-core`'s `DocumentModel`: several fields here
 * (keepNext, cantSplit, tblHeader, explicit page/column breaks, sections,
 * columns) are not yet represented in the shared model. `adapter.ts` maps
 * what the current model exposes and leaves the rest at their (Word-default)
 * fallback, so the pipeline keeps compiling as other agents add model
 * fields, and nothing here ever reads `(x as any)`.
 */
export interface LayoutRun {
	text: string;
	bold?: boolean;
	italic?: boolean;
	fontFamily?: string;
	/** Points, matching `TextRun.fontSize` in docx-core. */
	fontSizePt?: number;
	/** A hard page/column break placed immediately after this run's text. */
	breakAfter?: 'page' | 'column';
	/** An inline picture: occupies its size on the line (`text` is empty). */
	object?: LayoutObject;
}

/** A picture's package part and size in CSS pixels; the engine never loads its bytes. */
export interface LayoutObject {
	partName: string;
	contentType: string;
	widthPx: number;
	heightPx: number;
}

/** A floating picture anchored to a paragraph (`wp:anchor`), positioned on the page. */
export interface LayoutFloat extends LayoutObject {
	/** `wp:positionH/@relativeFrom`: page, margin, column, leftMargin, rightMargin, character… */
	relativeFromH?: string;
	alignH?: 'left' | 'center' | 'right' | 'inside' | 'outside';
	offsetXPx?: number;
	/** `wp:positionV/@relativeFrom`: page, margin, paragraph, line, topMargin, bottomMargin… */
	relativeFromV?: string;
	alignV?: 'top' | 'center' | 'bottom' | 'inside' | 'outside';
	offsetYPx?: number;
	behindText?: boolean;
	wrap?: string;
}

export type ParagraphAlign = 'left' | 'center' | 'right' | 'justify';
export type LineSpacingRule = 'auto' | 'exact' | 'atLeast';

export interface LayoutParagraph {
	kind: 'paragraph';
	id: string;
	runs: LayoutRun[];
	align?: ParagraphAlign;
	direction?: 'ltr' | 'rtl';
	spacingBeforeTwips?: number;
	spacingAfterTwips?: number;
	lineSpacingTwips?: number;
	lineSpacingRule?: LineSpacingRule;
	indentLeftTwips?: number;
	indentRightTwips?: number;
	indentStartTwips?: number;
	indentEndTwips?: number;
	firstLineTwips?: number;
	hangingTwips?: number;
	/** `w:contextualSpacing`: suppress spacing between paragraphs sharing `styleId`. */
	contextualSpacing?: boolean;
	styleId?: string;
	/** `w:pageBreakBefore`. */
	pageBreakBefore?: boolean;
	/** `w:keepNext`: keep this paragraph on the same page as the block after it. */
	keepNext?: boolean;
	/** `w:keepLines`: never split this paragraph's own lines across pages. */
	keepLines?: boolean;
	/** `w:widowControl`; Word's document default is on. Undefined means on. */
	widowControl?: boolean;
	/** Floating pictures anchored in this paragraph. */
	floats?: LayoutFloat[];
}

export interface LayoutTableCell {
	paragraphs: LayoutParagraph[];
	/** Column width in CSS pixels; undefined cells share the row's remaining width evenly. */
	widthPx?: number;
}

export interface LayoutTableRow {
	cells: LayoutTableCell[];
	/** `w:trPr/w:cantSplit`: this row must stay on one page. */
	cantSplit?: boolean;
	/** `w:trPr/w:tblHeader`: repeat this row at the top of every page the table continues onto. */
	isHeader?: boolean;
}

export interface LayoutTable {
	kind: 'table';
	id: string;
	rows: LayoutTableRow[];
}

export type LayoutBlock = LayoutParagraph | LayoutTable;

export interface LayoutPageGeometry {
	widthPx: number;
	heightPx: number;
	marginTopPx: number;
	marginRightPx: number;
	marginBottomPx: number;
	marginLeftPx: number;
}

export interface LayoutColumns {
	count: number;
	gapPx: number;
}

export interface LayoutSection {
	page: LayoutPageGeometry;
	columns?: LayoutColumns;
	blocks: LayoutBlock[];
	/** How this section begins relative to the previous one; the first section is always a fresh page. */
	break?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';
	/** `w:vAlign`: where content sits vertically on each page of the section. */
	verticalAlign?: 'top' | 'center' | 'both' | 'bottom';
}

export interface LayoutDocumentInput {
	sections: LayoutSection[];
}
