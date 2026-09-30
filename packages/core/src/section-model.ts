// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Sections, headers/footers and footnotes/endnotes.
import type { Block } from './model.js';
import type { StNumberFormat, StSectionMark, StVerticalJc } from './generated/wml-simple-types.js';
import type { SignedTwips, Twips } from './units.js';

/** Header/footer content, parsed with the same paragraph/table parser as the body. */
export interface HeaderFooterContent {
	blocks: Block[];
	/** The header or footer package part, e.g. `word/header1.xml`; several sections may share one. */
	partName?: string;
	/** Source package part to preserve XML and relationships when making an independent copy. */
	sourcePartName?: string;
}
/** One `w:headerReference`/`w:footerReference` slot resolved into content, if present. */
export interface HeaderFooterSlots {
	default?: HeaderFooterContent;
	even?: HeaderFooterContent;
	first?: HeaderFooterContent;
}
export interface SectionColumn {
	widthTwips: Twips;
	spacingTwips?: Twips;
}
/** `w:cols`: newspaper-style column layout for the section. */
export interface SectionColumns {
	count: number;
	spacingTwips?: Twips;
	equalWidth: boolean;
	separator?: boolean;
	/** Present only when `equalWidth` is false and individual `w:col` widths were given. */
	widths?: SectionColumn[];
}
export interface SectionPageNumbering {
	start?: number;
	/** Raw `w:pgNumType/@w:fmt` token, e.g. `decimal`, `upperRoman`. */
	format?: StNumberFormat;
}
/**
 * One `w:sectPr` (paragraph-level section break or the final body section), in native Word
 * units. The pagination engine is the primary consumer; this package does not lay out pages.
 */
export interface LineNumberSettings {
	/** Show a number on every n-th line (`w:countBy`, at least 1). */
	countBy: number;
	/** First visible number of a restart (one more than the zero-based `w:start`, default 1). */
	start: number;
	/** `w:restart`: start over each page, each section, or never. */
	restart: 'newPage' | 'newSection' | 'continuous';
	/** Gap between the number and the text (`w:distance`); undefined lets the renderer choose. */
	distanceTwips?: number;
}
export interface SectionProperties {
	/** Id of the last block (paragraph or table) this section covers. */
	endsAtBlockId: string;
	type: StSectionMark;
	pageWidthTwips: Twips;
	pageHeightTwips: Twips;
	orientation: 'portrait' | 'landscape';
	marginTopTwips: SignedTwips;
	marginRightTwips: Twips;
	marginBottomTwips: SignedTwips;
	marginLeftTwips: Twips;
	headerDistanceTwips?: Twips;
	footerDistanceTwips?: Twips;
	gutterTwips?: Twips;
	columns: SectionColumns;
	/** `w:titlePg`: the section's first page uses distinct first-page headers/footers. */
	titlePage?: boolean;
	verticalAlign?: StVerticalJc;
	pageNumbering?: SectionPageNumbering;
	/** `w:lnNumType` is present (line numbers are on for this section). */
	lineNumbering?: boolean;
	/** The `w:lnNumType` values behind `lineNumbering`. */
	lineNumberSettings?: LineNumberSettings;
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
