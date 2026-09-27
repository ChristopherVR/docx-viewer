// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { NumberingCatalog } from './numbering-model.js';
export interface TextRun {
	text: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	/** Word named highlight color, such as `yellow` or `lightGray`. */
	highlight?: string;
	verticalAlign?: 'superscript' | 'subscript';
	/** Direct Word run language tag (`w:lang/@w:val`), without automatic detection. */
	language?: string;
	/** Direct East Asian script language tag (`w:lang/@w:eastAsia`). */
	eastAsiaLanguage?: string;
	/** Direct complex-script language tag (`w:lang/@w:bidi`). */
	bidiLanguage?: string;
	/** Explicit run-level bidirectional override; `false` means direct off, undefined inherits. */
	rtl?: boolean;
	/** Font size in points. */
	fontSize?: number;
	fontFamily?: string;
	/** Direct RGB color, e.g. #28665E. Theme resolution is not yet supported. */
	color?: string;
	/**
	 * A run that is itself a page or column break marker (`w:br` type page/column) instead of
	 * visible text. `text` is empty on these runs; editors render a distinct visible marker.
	 */
	break?: 'page' | 'column';
	/**
	 * A run that is itself a footnote/endnote reference mark (`w:footnoteReference` /
	 * `w:endnoteReference`) instead of visible text. `text` is empty on these runs; the numeric
	 * mark is derived from document order, not stored here.
	 */
	noteReference?: { kind: 'footnote' | 'endnote'; id: string };
}
export interface Paragraph {
	type: 'paragraph';
	/** Stable identity across edits, unique throughout the document. */
	id: string;
	runs: TextRun[];
	align?: 'left' | 'center' | 'right' | 'justify';
	/** Paragraph base direction from direct `w:bidi`; undefined inherits. */
	direction?: 'ltr' | 'rtl';
	style?: string;
	/** Word paragraph spacing and indentation values, kept in their native twip units (1/20 pt). */
	spacingBeforeTwips?: number;
	spacingAfterTwips?: number;
	lineSpacingTwips?: number;
	/** `auto` uses 240ths of a line; `exact` and `atLeast` use twips. */
	lineSpacingRule?: 'auto' | 'exact' | 'atLeast';
	indentLeftTwips?: number;
	indentRightTwips?: number;
	indentStartTwips?: number;
	indentEndTwips?: number;
	firstLineTwips?: number;
	hangingTwips?: number;
	/** Direct `w:numPr` on this paragraph; undefined may still inherit numbering through `style`. */
	numbering?: { numId: number; level: number };
	/** `w:pPr/w:pageBreakBefore`: forces this paragraph to start a new page. */
	pageBreakBefore?: boolean;
}
/** Direct paragraph properties supported by the editor, in native Word units. */
export type ParagraphFormatting = Pick<
	Paragraph,
	| 'align'
	| 'direction'
	| 'spacingBeforeTwips'
	| 'spacingAfterTwips'
	| 'lineSpacingTwips'
	| 'lineSpacingRule'
	| 'indentLeftTwips'
	| 'indentRightTwips'
	| 'indentStartTwips'
	| 'indentEndTwips'
	| 'firstLineTwips'
	| 'hangingTwips'
>;
export interface ParagraphStyleDefinition {
	id: string;
	name?: string;
	basedOn?: string;
	isDefault?: boolean;
	formatting: ParagraphFormatting;
	/** Numbering inherited by paragraphs using this style through `w:pPr/w:numPr` in `styles.xml`. */
	numbering?: { numId: number; level: number };
}
/** Parsed source style data. It is retained for inheritance and never flattened on save. */
export interface ParagraphStyleCatalog {
	docDefaults: ParagraphFormatting;
	styles: Record<string, ParagraphStyleDefinition>;
	warnings: string[];
}
export interface TableCell {
	paragraphs: Paragraph[];
}
export interface Table {
	type: 'table';
	id: string;
	/** Whether row, cell and paragraph structure can be edited without losing unsupported table XML. */
	structureEditable?: boolean;
	rows: TableCell[][];
}
export type Block = Paragraph | Table;
/** Read-only header/footer content, parsed with the same paragraph/table parser as the body. */
export interface HeaderFooterContent {
	blocks: Block[];
}
/** One `w:headerReference`/`w:footerReference` slot resolved into read-only content, if present. */
export interface HeaderFooterSlots {
	default?: HeaderFooterContent;
	even?: HeaderFooterContent;
	first?: HeaderFooterContent;
}
export interface SectionColumn {
	widthTwips: number;
	spacingTwips?: number;
}
/** `w:cols`: newspaper-style column layout for the section. */
export interface SectionColumns {
	count: number;
	spacingTwips?: number;
	equalWidth: boolean;
	separator?: boolean;
	/** Present only when `equalWidth` is false and individual `w:col` widths were given. */
	widths?: SectionColumn[];
}
export interface SectionPageNumbering {
	start?: number;
	/** Raw `w:pgNumType/@w:fmt` token, e.g. `decimal`, `upperRoman`. */
	format?: string;
}
/**
 * One `w:sectPr` (paragraph-level section break or the final body section), in native Word
 * units. The pagination engine is the primary consumer; this package does not lay out pages.
 */
export interface SectionProperties {
	/** Id of the last block (paragraph or table) this section covers. */
	endsAtBlockId: string;
	type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
	pageWidthTwips: number;
	pageHeightTwips: number;
	orientation: 'portrait' | 'landscape';
	marginTopTwips: number;
	marginRightTwips: number;
	marginBottomTwips: number;
	marginLeftTwips: number;
	headerDistanceTwips?: number;
	footerDistanceTwips?: number;
	gutterTwips?: number;
	columns: SectionColumns;
	/** `w:titlePg`: the section's first page uses distinct first-page headers/footers. */
	titlePage?: boolean;
	verticalAlign?: 'top' | 'center' | 'both' | 'bottom';
	pageNumbering?: SectionPageNumbering;
	/** `w:lnNumType` presence; line numbering values themselves are not modeled. */
	lineNumbering?: boolean;
	/** `w:pgBorders` presence; border styling itself is not modeled. */
	pageBorders?: boolean;
	headers?: HeaderFooterSlots;
	footers?: HeaderFooterSlots;
}
/** One footnote or endnote body, parsed with the same paragraph/table parser as the body. */
export interface Note {
	id: string;
	blocks: Block[];
}
export interface DocumentModel {
	blocks: Block[];
	/** CSS pixels at 96 DPI. Page height is currently a surface minimum, not pagination. */
	page: {
		width: number;
		height: number;
		marginTop: number;
		marginRight: number;
		marginBottom: number;
		marginLeft: number;
	};
	warnings: string[];
	/** Source paragraph defaults/styles; editing the catalog itself is not supported. */
	paragraphStyles?: ParagraphStyleCatalog;
	/** Parsed `word/numbering.xml`. Existing entries are read-only; new list definitions may be added. */
	numberingCatalog?: NumberingCatalog;
	/**
	 * Every parsed section in document order (paragraph-level section breaks, then the final
	 * body section). `page` above mirrors the last section's page geometry for backward compat.
	 */
	sections?: SectionProperties[];
	/** Whether `settings.xml` requests separate even-page headers/footers (`w:evenAndOddHeaders`). */
	evenAndOddHeaders?: boolean;
	footnotes?: Note[];
	endnotes?: Note[];
	/** Raw `w:footnotePr/w:numFmt` token from settings.xml; defaults to `decimal`. */
	footnoteNumFmt?: string;
	/** Raw `w:endnotePr/w:numFmt` token from settings.xml; defaults to `lowerRoman`. */
	endnoteNumFmt?: string;
}
export interface LoadedDocument {
	model: DocumentModel;
	/** Returns bytes in the original format; unsupported edits reject instead of silently degrading it. */
	save(model?: DocumentModel): Promise<Uint8Array>;
}
export function createDocument(): DocumentModel {
	return {
		blocks: [{ type: 'paragraph', id: 'p1', runs: [{ text: '' }] }],
		page: {
			width: 816,
			height: 1056,
			marginTop: 96,
			marginRight: 96,
			marginBottom: 96,
			marginLeft: 96,
		},
		warnings: [],
	};
}
