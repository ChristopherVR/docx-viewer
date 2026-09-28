/** One laid-out fragment of text within a line, styled by a single run. */
export interface LayoutFragment {
	text: string;
	xPx: number;
	widthPx: number;
	runIndex: number;
	bold?: boolean;
	italic?: boolean;
	fontFamily?: string;
	fontSizePt?: number;
	/** An inline picture drawn in this fragment's box, bottom-aligned on the line. */
	object?: import('./input.js').LayoutObject;
	color?: string;
	underline?: boolean;
	strike?: boolean;
	/** A tab's leader fill (dots, hyphens, a line…) drawn across its width. */
	leader?: 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

/** One visual line of a paragraph, positioned within its block's box. */
export interface LayoutLine {
	yPx: number;
	heightPx: number;
	fragments: LayoutFragment[];
	/** Space skipped above this line to clear a picture wrapped top-and-bottom; `yPx` is below it. */
	gapBeforePx?: number;
	/** Character range within the paragraph's concatenated run text, for hit-testing. */
	sourceStart: number;
	sourceEnd: number;
}

export interface LayoutParagraphBox {
	kind: 'paragraph';
	blockId: string;
	yPx: number;
	heightPx: number;
	lines: LayoutLine[];
}

/** Where a cell sits in its row and how it is drawn (borders, shading, padding). */
export interface LayoutCellGeometry {
	xPx: number;
	widthPx: number;
	paddingLeftPx: number;
	paddingRightPx: number;
	borders?: import('./input.js').LayoutCellBorders;
	shading?: string;
	verticalAlign?: 'top' | 'center' | 'bottom';
	/** Height of the cell's content (with its top and bottom padding), for vertical alignment. */
	contentHeightPx: number;
}

export interface LayoutTableRowBox {
	yPx: number;
	heightPx: number;
	/** True when this row is a repeated header copy rather than the row's first placement. */
	repeated: boolean;
	cells: LayoutParagraphBox[][];
	/** Per-cell position and appearance, parallel to `cells`. */
	geometry?: LayoutCellGeometry[];
}

export interface LayoutTableBox {
	kind: 'table';
	blockId: string;
	/** Offset from the column's left edge (table indent or alignment). */
	xPx?: number;
	yPx: number;
	heightPx: number;
	rows: LayoutTableRowBox[];
}

export type LayoutBlockBox = LayoutParagraphBox | LayoutTableBox;

export interface LayoutColumnBox {
	xPx: number;
	widthPx: number;
	blocks: LayoutBlockBox[];
}

export interface LayoutPageBox {
	index: number;
	/** Index into the input sections this page belongs to. */
	sectionIndex: number;
	/** Zero-based position of this page within its section (for first-page headers and numbering). */
	pageInSection: number;
	widthPx: number;
	heightPx: number;
	marginTopPx: number;
	marginRightPx: number;
	marginBottomPx: number;
	marginLeftPx: number;
	columns: LayoutColumnBox[];
	/** Floating pictures on this page, in page coordinates (CSS pixels from the sheet's top-left). */
	floats?: LayoutFloatBox[];
}

export interface LayoutFloatBox {
	blockId: string;
	xPx: number;
	yPx: number;
	widthPx: number;
	heightPx: number;
	partName: string;
	contentType: string;
	behindText: boolean;
	/** `wp:wrapSquare`, `wrapTopAndBottom`… (`none` floats over or behind the text). */
	wrap?: string;
}

export interface LayoutResult {
	pages: LayoutPageBox[];
	/** Honest, human-readable notes on approximations relative to Word (see docs/parity-roadmap.md §3). */
	approximations: string[];
}
