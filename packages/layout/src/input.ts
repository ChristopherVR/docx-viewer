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
	/** Generated text (a list label) that is not part of the paragraph's own characters. */
	synthetic?: boolean;
	/** `#rrggbb`, display only. */
	color?: string;
	underline?: boolean;
	strike?: boolean;
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

/** A paragraph tab stop, positioned from the paragraph's text margin (not its indent). */
export interface LayoutTabStop {
	posPx: number;
	align: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'start' | 'end' | 'num';
	leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
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
	/** Custom tab stops (`w:tabs`). */
	tabStops?: LayoutTabStop[];
}

/** A resolved cell border, ready to draw. */
export interface LayoutBorder {
	widthPx: number;
	style: 'solid' | 'double' | 'dotted' | 'dashed';
	color: string;
}
export interface LayoutCellBorders {
	top?: LayoutBorder;
	right?: LayoutBorder;
	bottom?: LayoutBorder;
	left?: LayoutBorder;
}
export interface LayoutCellPadding {
	top: number;
	right: number;
	bottom: number;
	left: number;
}

export interface LayoutTableCell {
	paragraphs: LayoutParagraph[];
	/** Column width in CSS pixels; undefined cells share the row's remaining width evenly. */
	widthPx?: number;
	/** Left edge from the table's left, from the table grid (merged cells span grid columns). */
	xPx?: number;
	/** Cell margins (`w:tcMar`, else the table default); content is inset by them. */
	padding?: LayoutCellPadding;
	borders?: LayoutCellBorders;
	/** `#rrggbb` shading fill. */
	shading?: string;
	verticalAlign?: 'top' | 'center' | 'bottom';
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
	/** Total grid width, for aligning the table in its column. */
	widthPx?: number;
	/** `w:tblInd`: offset from the column's left edge. */
	indentPx?: number;
	/** `w:jc` on the table. */
	alignment?: 'left' | 'center' | 'right';
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
