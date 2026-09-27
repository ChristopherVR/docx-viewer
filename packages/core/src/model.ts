// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
/** A tracked-change revision recorded on a run or paragraph mark. */
export interface Revision {
	/** `moveFrom`/`moveTo` are tracked as delete/insert pairs; Word's move linkage is not modeled. */
	kind: 'insert' | 'delete' | 'moveFrom' | 'moveTo' | 'formatChange' | 'paragraphChange';
	author: string;
	date?: string;
	/** Source `w:id`; not guaranteed unique outside the paragraph it was parsed from. */
	id: string;
}
/** A comment thread entry parsed from comments.xml / commentsExtended.xml. */
export interface Comment {
	id: string;
	author: string;
	initials?: string;
	date?: string;
	text: string;
	/** commentsExtended.xml `w15:done`. */
	resolved?: boolean;
	/** commentsExtended.xml parent linkage for threaded replies. */
	parentId?: string;
}
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
	/** Tracked-change metadata for this run; absent means the run has no pending revision. */
	revision?: Revision;
	/** IDs of comments whose range covers this run. */
	commentIds?: string[];
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
	/** Tracked insertion/deletion of the paragraph mark itself (the paragraph break). */
	markRevision?: Revision;
	/** Marks that `w:pPrChange` recorded a prior paragraph formatting snapshot; the snapshot itself is not modeled. */
	formatRevision?: Revision;
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
	/** Comment threads parsed from comments.xml / commentsExtended.xml. */
	comments?: Comment[];
	/** settings.xml `w:trackRevisions`; toggling this changes how the editor records new edits. */
	trackChanges?: boolean;
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
