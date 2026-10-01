// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Paragraph } from './model.js';
import type { RunFormatting } from './run-style-model.js';
import type { ThemeColorReference, ThemeColorToken } from './theme-model.js';
import type { StBorder, StVerticalJc } from './generated/wml-simple-types.js';
import type { EighthPoints, SignedTwips, Twips } from './units.js';

export interface TableBorderSide {
	style?: StBorder;
	/** Border width in eighths of a point (`w:sz`). */
	sizeEighthPoints?: EighthPoints;
	color?: string;
	themeColor?: ThemeColorToken;
	/** Paragraph borders only: gap between the border and the text (`w:space`), in points. */
	spacePoints?: number;
	/** Page borders only: a Word art border (`w:art`); drawn as a plain line. */
	art?: string;
}
/** Paragraph borders (`w:pBdr`); `between` separates paragraphs that share the same borders. */
export interface ParagraphBorders {
	top?: TableBorderSide;
	bottom?: TableBorderSide;
	left?: TableBorderSide;
	right?: TableBorderSide;
	between?: TableBorderSide;
}
export interface TableBorders {
	top?: TableBorderSide;
	bottom?: TableBorderSide;
	left?: TableBorderSide;
	right?: TableBorderSide;
	insideH?: TableBorderSide;
	insideV?: TableBorderSide;
}
export interface TableCellMargins {
	top?: SignedTwips;
	bottom?: SignedTwips;
	left?: SignedTwips;
	right?: SignedTwips;
}
/** Row-level properties (`w:trPr`): height, keep-together and repeat-as-header. */
export interface TableRowProperties {
	heightTwips?: Twips;
	/** `exact` fixes the height; `atLeast` (Word's `atLeast` and `auto`) is a minimum. */
	heightRule?: 'atLeast' | 'exact';
	/** `w:cantSplit`: the row is never split across pages. */
	cantSplit?: boolean;
	/** `w:tblHeader`: the row repeats at the top of each page the table continues onto. */
	header?: boolean;
}
/** A read-only preview of a nested table's text, rendered but not independently editable. */
export interface NestedTablePreview {
	rows: { text: string }[][];
}
export interface TableCell {
	paragraphs: Paragraph[];
	/** `w:gridSpan`; number of grid columns this cell occupies (horizontal merge). */
	gridSpan?: number;
	/** `w:vMerge`; `restart` begins a vertical merge, `continue` extends the cell above. */
	verticalMerge?: 'restart' | 'continue';
	widthTwips?: Twips;
	verticalAlign?: StVerticalJc;
	shadingFill?: string;
	shadingThemeFill?: ThemeColorReference;
	borders?: TableBorders;
	margins?: TableCellMargins;
	/** Tables nested directly in this cell's XML, preserved but read-only in the editor. */
	nestedTables?: NestedTablePreview[];
}
/** Which `tblStyle` conditional formatting regions apply, from `w:tblLook`. */
export interface TableLook {
	firstRow?: boolean;
	lastRow?: boolean;
	firstColumn?: boolean;
	lastColumn?: boolean;
	noHBand?: boolean;
	noVBand?: boolean;
}
export type TableConditionalRegion =
	| 'wholeTable'
	| 'firstRow'
	| 'lastRow'
	| 'firstCol'
	| 'lastCol'
	| 'band1Horz'
	| 'band2Horz'
	| 'band1Vert'
	| 'band2Vert';
/** A conditional-formatting region (`w:tblStylePr`) within a table style. */
export interface TableStyleConditionalFormatting {
	borders?: TableBorders;
	shadingFill?: string;
	shadingThemeFill?: ThemeColorReference;
	run?: RunFormatting;
}
export interface TableStyleDefinition {
	id: string;
	name?: string;
	basedOn?: string;
	isDefault?: boolean;
	borders?: TableBorders;
	shadingFill?: string;
	conditional: Partial<Record<TableConditionalRegion, TableStyleConditionalFormatting>>;
}
/** Parsed `styles.xml` table styles; conditional formatting resolves for rendering only. */
export interface TableStyleCatalog {
	styles: Record<string, TableStyleDefinition>;
	warnings: string[];
}
