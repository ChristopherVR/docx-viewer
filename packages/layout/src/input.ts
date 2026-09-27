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
}

export interface LayoutDocumentInput {
	sections: LayoutSection[];
}
