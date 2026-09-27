// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { NumberingCatalog } from './numbering-model.js';
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
	/**
	 * Present on runs holding a field's displayed result (`w:fldSimple`, or text between a complex
	 * field's `separate` and `end`). Display metadata only: fields are not recalculated on save and
	 * paragraphs containing them stay protected from edits.
	 */
	field?: { instr: string };
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
	/** Direct `w:numPr` on this paragraph; undefined may still inherit numbering through `style`. */
	numbering?: { numId: number; level: number };
	/** `w:pPr/w:pageBreakBefore`: forces this paragraph to start a new page. */
	pageBreakBefore?: boolean;
	/** Tracked insertion/deletion of the paragraph mark itself (the paragraph break). */
	markRevision?: Revision;
	/** Marks that `w:pPrChange` recorded a prior paragraph formatting snapshot; the snapshot itself is not modeled. */
	formatRevision?: Revision;
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
}
export type Block = Paragraph | Table;
/** Read-only header/footer content, parsed with the same paragraph/table parser as the body. */
export interface HeaderFooterContent {
	blocks: Block[];
	/** The header or footer package part, e.g. `word/header1.xml`; several sections may share one. */
	partName?: string;
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
	};
}
