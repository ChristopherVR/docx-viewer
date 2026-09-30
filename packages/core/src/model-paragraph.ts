// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { StJc, StTabJc, StTabTlc } from './generated/wml-simple-types.js';
import type { Revision, TextRun } from './model-run.js';
import type { SignedTwips, Twips } from './units.js';

/** A paragraph tab stop (`w:tab`): position from the text margin, alignment and leader fill. */
export interface TabStop {
	posTwips: SignedTwips;
	align: StTabJc;
	leader?: StTabTlc;
}
export interface Paragraph {
	type: 'paragraph';
	/** Stable identity across edits, unique throughout the document. */
	id: string;
	runs: TextRun[];
	align?: 'left' | 'center' | 'right' | 'justify';
	/**
	 * The exact `w:jc` value read from the file (`start`, `end`, `distribute`, kashida variants...).
	 * `align` is derived from it for rendering; it is only authoritative while `align` still equals
	 * `alignFromJustification(justification, rtl)`, since editors change `align` alone.
	 */
	justification?: StJc;
	/** Paragraph base direction from direct `w:bidi`; undefined inherits. */
	direction?: 'ltr' | 'rtl';
	style?: string;
	/** Word paragraph spacing and indentation values, kept in their native twip units (1/20 pt). */
	spacingBeforeTwips?: Twips;
	spacingAfterTwips?: Twips;
	lineSpacingTwips?: SignedTwips;
	/** `auto` uses 240ths of a line; `exact` and `atLeast` use twips. */
	lineSpacingRule?: 'auto' | 'exact' | 'atLeast';
	indentLeftTwips?: SignedTwips;
	indentRightTwips?: SignedTwips;
	indentStartTwips?: SignedTwips;
	indentEndTwips?: SignedTwips;
	firstLineTwips?: Twips;
	hangingTwips?: Twips;
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
	/** `w:keepNext`: keep this paragraph on the same page as the next. `false` cancels a style. */
	keepNext?: boolean;
	/** `w:keepLines`: never split this paragraph across pages. */
	keepLines?: boolean;
	/** `w:widowControl`: avoid single lines at a page's top or bottom (on unless turned off). */
	widowControl?: boolean;
	/** `w:contextualSpacing`: no spacing between paragraphs of the same style. */
	contextualSpacing?: boolean;
	/**
	 * `w:framePr/@w:dropCap`: this paragraph is a drop cap frame (the initial letter) for the
	 * paragraph that follows. `lines` is the number of text lines it spans.
	 */
	dropCap?: { style: 'drop' | 'margin'; lines: number };
	/** `w:pBdr`; read-only (preserved in the source XML, not written for new paragraphs). */
	borders?: import('./table-model.js').ParagraphBorders;
	/** `w:shd/@w:fill` as `#rrggbb`; read-only like `borders`. */
	shadingFill?: string;
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
	| 'justification'
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
	| 'keepNext'
	| 'keepLines'
	| 'widowControl'
	| 'contextualSpacing'
	| 'borders'
	| 'shadingFill'
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
