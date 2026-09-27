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
}

/** One visual line of a paragraph, positioned within its block's box. */
export interface LayoutLine {
	yPx: number;
	heightPx: number;
	fragments: LayoutFragment[];
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

export interface LayoutTableRowBox {
	yPx: number;
	heightPx: number;
	/** True when this row is a repeated header copy rather than the row's first placement. */
	repeated: boolean;
	cells: LayoutParagraphBox[][];
}

export interface LayoutTableBox {
	kind: 'table';
	blockId: string;
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
	widthPx: number;
	heightPx: number;
	marginTopPx: number;
	marginRightPx: number;
	marginBottomPx: number;
	marginLeftPx: number;
	columns: LayoutColumnBox[];
}

export interface LayoutResult {
	pages: LayoutPageBox[];
	/** Honest, human-readable notes on approximations relative to Word (see docs/parity-roadmap.md §3). */
	approximations: string[];
}
