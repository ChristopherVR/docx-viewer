// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Sections, headers/footers and footnotes/endnotes.
import type { Block } from './model.js';

/** Header/footer content, parsed with the same paragraph/table parser as the body. */
export interface HeaderFooterContent {
	blocks: Block[];
	/** The header or footer package part, e.g. `word/header1.xml`; several sections may share one. */
	partName?: string;
}
/** One `w:headerReference`/`w:footerReference` slot resolved into content, if present. */
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
