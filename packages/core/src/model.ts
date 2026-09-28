// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { defaultStyleCatalogs } from './default-styles.js';
import type { NumberingCatalog } from './numbering-model.js';
import type { HyperlinkInfo, InlineImage } from './inline-model.js';
export type { HyperlinkInfo, InlineImage, PicturePlacement } from './inline-model.js';
import type { HeaderFooterContent, Note, SectionProperties } from './section-model.js';
export type {
	HeaderFooterContent,
	HeaderFooterSlots,
	Note,
	SectionColumn,
	SectionColumns,
	SectionPageNumbering,
	SectionProperties,
} from './section-model.js';
/** A tracked-change revision recorded on a run or paragraph mark. */
export interface Revision {
	kind: 'insert' | 'delete' | 'moveFrom' | 'moveTo' | 'formatChange' | 'paragraphChange';
	/**
	 * For `moveFrom`/`moveTo`: the move this text belongs to. Both sides share `name` (from
	 * `w:moveFromRangeStart`/`w:moveToRangeStart`); `rangeId` is that range marker's `w:id`.
	 */
	move?: { name: string; rangeId?: string };
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
import type {
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
	/** Direct RGB color, e.g. #28665E. May coexist with `colorTheme` as Word's stored fallback. */
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
	/** The automatic number mark (`w:footnoteRef`/`w:endnoteRef`) that starts a note's own text. */
	noteMark?: 'footnote' | 'endnote';
	/**
	 * Present on runs holding a field's displayed result (`w:fldSimple`, or text between a complex
	 * field's `separate` and `end`). Display metadata only: fields are not recalculated on save and
	 * paragraphs containing them stay protected from edits.
	 */
	field?: { instr: string; simple?: boolean };
	/** A complex field's `w:fldChar` marker run (begin, separate or end); `text` is empty. */
	fieldChar?: 'begin' | 'separate' | 'end';
	/** A complex field's instruction text run (`w:instrText`), e.g. ` TOC \o "1-3" `; `text` is empty. */
	fieldCode?: string;
	/** Tracked-change metadata for this run; absent means the run has no pending revision. */
	revision?: Revision;
	/** IDs of comments whose range covers this run. */
	commentIds?: string[];
	/** Direct `w:color/@w:themeColor` (+ themeTint/themeShade); resolution happens in a separate layer. */
	colorTheme?: ThemeColorReference;
	/** Character style reference (`w:rStyle/@w:val`); preserved and editable, not flattened. */
	style?: string;
	caps?: boolean;
	smallCaps?: boolean;
	/** `w:dstrike`; kept distinct from the single-line `strike` toggle. */
	doubleStrike?: boolean;
	/** Hidden text (`w:vanish`); the editor renders it dimmed rather than removing it. */
	vanish?: boolean;
	/** Non-single underline style, e.g. `double`/`wave`; `underline` stays the simple on/off toggle. */
	underlineStyle?: WordUnderlineStyle;
	underlineColor?: string;
	/** `w:spacing/@w:val` character spacing, in twips (positive expands, negative condenses). */
	characterSpacingTwips?: number;
	/** Direct `w:shd/@w:fill` run shading. */
	shadingFill?: string;
	/** Direct `w:shd` theme fill; kept alongside `shadingFill` without flattening. */
	shadingThemeFill?: ThemeColorReference;
	/** Direct `w:rFonts` theme font references, per script; resolved via the document theme. */
	fontTheme?: Partial<Record<ThemeFontScript, ThemeFontRole>>;
	/** Present when this run is an inline picture instead of text; `text` is empty. */
	image?: InlineImage;
	/** Hyperlink target for this run, from `w:hyperlink` (or a simple `HYPERLINK` field). */
	link?: HyperlinkInfo;
}
/** A paragraph tab stop (`w:tab`): position from the text margin, alignment and leader fill. */
export interface TabStop {
	posTwips: number;
	align: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'start' | 'end' | 'num';
	leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
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
	/** Tracked insertion/deletion of the paragraph mark itself (the paragraph break). */
	markRevision?: Revision;
	/** Marks that `w:pPrChange` recorded a prior paragraph formatting snapshot; the snapshot itself is not modeled. */
	formatRevision?: Revision;
	/** Custom tab stops (`w:tabs`), in document order. */
	tabStops?: TabStop[];
	/**
	 * Bookmark names starting in this paragraph (`w:bookmarkStart/@w:name`). Added names are
	 * written around the whole paragraph; removed names lose their markers here. Exact character
	 * ranges are not modeled.
	 */
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
	/** Numbering inherited by paragraphs using this style through `w:pPr/w:numPr` in `styles.xml`. */
	numbering?: { numId: number; level: number };
}
/** Parsed source style data. It is retained for inheritance and never flattened on save. */
export interface ParagraphStyleCatalog {
	docDefaults: ParagraphFormatting;
	styles: Record<string, ParagraphStyleDefinition>;
	warnings: string[];
}
export type TableCell = TableCellShape;
export interface Table {
	type: 'table';
	id: string;
	/** Whether row, cell and paragraph structure can be edited without losing unsupported table XML. */
	structureEditable?: boolean;
	rows: TableCell[][];
	/** `w:tblGrid/w:gridCol` widths in twips, one per grid column. */
	grid?: number[];
	widthTwips?: number;
	alignment?: 'left' | 'center' | 'right';
	indentTwips?: number;
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
	/** Whether `settings.xml` requests separate even-page headers/footers (`w:evenAndOddHeaders`). */
	evenAndOddHeaders?: boolean;
	footnotes?: Note[];
	endnotes?: Note[];
	/** Raw `w:footnotePr/w:numFmt` token from settings.xml; defaults to `decimal`. */
	footnoteNumFmt?: string;
	/** Raw `w:endnotePr/w:numFmt` token from settings.xml; defaults to `lowerRoman`. */
	endnoteNumFmt?: string;
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
