// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
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
	/** Present when this run is an inline picture instead of text; `text` is empty. */
	image?: InlineImage;
	/** Hyperlink target for this run, from `w:hyperlink` (or a simple `HYPERLINK` field). */
	link?: HyperlinkInfo;
}
/** An inline drawing (`w:drawing` or legacy `w:pict`) modeled at run granularity. */
export interface InlineImage {
	/** Relationship id in `word/_rels/document.xml.rels` pointing at the media part. */
	relId: string;
	/** Package part name holding the image bytes, e.g. `word/media/image1.png`. */
	partName: string;
	contentType: string;
	widthPx: number;
	heightPx: number;
	/** From `wp:docPr/@descr`. */
	altText?: string;
	/** From `wp:docPr/@title` or `@name`. */
	title?: string;
	/** `wp:anchor` (floating) drawings render as sized inline placeholders; wrapping/position is lost. */
	anchored?: boolean;
	/** Set for non-picture drawings (chart, SmartArt, shape, unresolved legacy VML): rendered as a labeled placeholder with no editable bytes. */
	unsupported?: string;
}
/** A `w:hyperlink` target, resolved from its relationship (external) or `w:anchor` (internal). */
export interface HyperlinkInfo {
	href?: string;
	anchor?: string;
	tooltip?: string;
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
	/** Read-only bookmark names starting in this paragraph (`w:bookmarkStart/@w:name`); bookmarks cannot be created or moved through the model. */
	bookmarks?: string[];
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
}
/** Bytes for a media part staged for save but not yet part of the loaded package (e.g. a newly inserted picture). */
export interface PendingMediaPart {
	bytes: Uint8Array;
	contentType: string;
}
export interface LoadedDocument {
	model: DocumentModel;
	/** Original image/media bytes keyed by package part name, kept off the JSON model so diffs stay small.
	 *  Absent (equivalent to empty) for loaders that never model inline media, such as legacy DOC. */
	media?: ReadonlyMap<string, Uint8Array>;
	/** Returns bytes in the original format; unsupported edits reject instead of silently degrading it.
	 *  `pendingMedia` supplies bytes for any `InlineImage.partName` newly referenced by `model` (e.g. inserted pictures). */
	save(
		model?: DocumentModel,
		pendingMedia?: ReadonlyMap<string, PendingMediaPart>,
	): Promise<Uint8Array>;
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
