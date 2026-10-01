import type { DocumentProperties } from './core-properties.js';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { defaultStyleCatalogs } from './default-styles.js';
import type { StJcTable, StNumberFormat } from './generated/wml-simple-types.js';
import type { ParagraphStyleCatalog } from './model-paragraph.js';
import type { NumberingCatalog } from './numbering-model.js';
import type { SignedTwips, Twips } from './units.js';
export type { Revision, TextRun } from './model-run.js';
export type {
	Paragraph,
	ParagraphFormatting,
	ParagraphStyleCatalog,
	ParagraphStyleDefinition,
	TabStop,
} from './model-paragraph.js';
import type { Paragraph } from './model-paragraph.js';
export type { HyperlinkInfo, InlineImage, PicturePlacement } from './inline-model.js';
import type { Note, SectionProperties } from './section-model.js';
export type {
	HeaderFooterContent,
	HeaderFooterSlots,
	LineNumberSettings,
	Note,
	PageBorders,
	PageBorderSide,
	SectionColumn,
	SectionColumns,
	SectionPageNumbering,
	SectionProperties,
} from './section-model.js';
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
export type {
	ThemeColorSlot,
	ThemeColorToken,
	ThemeFontSet,
	ThemeCatalog,
	ThemeColorReference,
	ThemeFontScript,
	ThemeFontRole,
	WordUnderlineStyle,
} from './theme-model.js';
export type {
	RunFormatting,
	CharacterStyleDefinition,
	RunStyleCatalog,
} from './run-style-model.js';
import type { RunStyleCatalog } from './run-style-model.js';
export type {
	TableBorderSide,
	TableBorders,
	TableCellMargins,
	NestedTablePreview,
	TableLook,
	TableConditionalRegion,
	TableStyleConditionalFormatting,
	TableStyleDefinition,
	TableStyleCatalog,
} from './table-model.js';
import type { TableCell as TableCellShape, TableStyleCatalog } from './table-model.js';

export type TableCell = TableCellShape;
export interface Table {
	type: 'table';
	id: string;
	/** Whether row, cell and paragraph structure can be edited without losing unsupported table XML. */
	structureEditable?: boolean;
	rows: TableCell[][];
	/** `w:tblGrid/w:gridCol` widths in twips, one per grid column. */
	grid?: Twips[];
	widthTwips?: Twips;
	alignment?: 'left' | 'center' | 'right';
	/** The exact `w:tblPr/w:jc` value (`ST_JcTable`, including `start`/`end`); `alignment` is derived from it. */
	justification?: StJcTable;
	indentTwips?: SignedTwips;
	borders?: import('./table-model.js').TableBorders;
	/** `w:tblStyle/@w:val`; conditional formatting resolves through `tableStyles` without flattening. */
	style?: string;
	look?: import('./table-model.js').TableLook;
	/** Default cell margins (`w:tblCellMar`); cells' own `w:tcMar` override them. */
	cellMargins?: import('./table-model.js').TableCellMargins;
	/** Per-row properties, parallel to `rows` (read from `w:trPr`; preserved on save). */
	rowProperties?: import('./table-model.js').TableRowProperties[];
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
	/** Parsed `word/numbering.xml`. Existing entries are read-only; new list definitions may be added. */
	numberingCatalog?: NumberingCatalog;
	/**
	 * Every parsed section in document order (paragraph-level section breaks, then the final
	 * body section). `page` above mirrors the last section's page geometry for backward compat.
	 */
	sections?: SectionProperties[];
	/** Package properties from `docProps/core.xml` (title, author, tags...). */
	properties?: DocumentProperties;
	/** Whether `settings.xml` requests separate even-page headers/footers (`w:evenAndOddHeaders`). */
	evenAndOddHeaders?: boolean;
	/** `w:autoHyphenation` in settings.xml: Word hyphenates words at line ends. */
	autoHyphenation?: boolean;
	/** Page colour (`w:document/w:background`), six uppercase hex digits without `#`; unset for none. */
	pageColor?: string;
	footnotes?: Note[];
	endnotes?: Note[];
	/** Raw `w:footnotePr/w:numFmt` token from settings.xml; defaults to `decimal`. */
	footnoteNumFmt?: StNumberFormat;
	/** Raw `w:endnotePr/w:numFmt` token from settings.xml; defaults to `lowerRoman`. */
	endnoteNumFmt?: StNumberFormat;
	/** Comment threads parsed from comments.xml / commentsExtended.xml. */
	comments?: Comment[];
	/** settings.xml `w:trackRevisions`; toggling this changes how the editor records new edits. */
	trackChanges?: boolean;
	/** Source run-level defaults/character styles; editing the catalog itself is not supported. */
	characterStyles?: RunStyleCatalog;
	/** Source table style catalog; editing the catalog itself is not supported. */
	tableStyles?: TableStyleCatalog;
	/** Parsed word/theme/theme1.xml and settings.xml color scheme mapping. */
	theme?: import('./theme-model.js').ThemeCatalog;
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
		...defaultStyleCatalogs(),
	};
}
