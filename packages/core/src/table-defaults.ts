// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Properties for tables created by the editor: Word requires `w:tblPr` and `w:tblGrid` on every
// table, and "Insert Table" gives new tables single-line borders on every edge.
import type { Table } from './model.js';
import type { TableBorderSide, TableBorders } from './table-model.js';
import { makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';

const SINGLE: TableBorderSide = { style: 'single', sizeEighthPoints: 4 };

/** Word's default borders for a newly inserted table (single 1/2 pt lines, automatic color). */
export const DEFAULT_TABLE_BORDERS: TableBorders = {
	top: SINGLE,
	left: SINGLE,
	bottom: SINGLE,
	right: SINGLE,
	insideH: SINGLE,
	insideV: SINGLE,
};

const BORDER_ORDER = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;

function setW(element: XmlElement, name: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${name}`, value);
}

/** Serializes `borders` as `w:tblBorders` (or `w:tcBorders`) in schema order. */
export function buildBorders(
	doc: XmlDocument,
	borders: TableBorders,
	tag: 'tblBorders' | 'tcBorders',
): XmlElement {
	const element = makeW(doc, tag);
	for (const side of BORDER_ORDER) {
		const value = borders[side];
		if (!value) continue;
		const edge = makeW(doc, side);
		setW(edge, 'val', value.style ?? 'single');
		setW(edge, 'sz', String(value.sizeEighthPoints ?? 4));
		setW(edge, 'space', '0');
		setW(edge, 'color', value.color ? value.color.replace(/^#/, '').toUpperCase() : 'auto');
		if (value.themeColor) setW(edge, 'themeColor', value.themeColor);
		element.appendChild(edge);
	}
	return element;
}

/** Number of grid columns a table spans (the widest row, counting horizontal merges). */
export function tableColumnCount(table: Table): number {
	return Math.max(
		1,
		...table.rows.map((row) => row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0)),
	);
}

/**
 * Builds `w:tblPr` and `w:tblGrid` for a table the editor created. Column widths come from the
 * model's grid when present, otherwise the content width is shared equally.
 */
export function buildNewTableProperties(
	doc: XmlDocument,
	table: Table,
	contentWidthTwips: number,
): { tblPr: XmlElement; tblGrid: XmlElement } {
	const tblPr = makeW(doc, 'tblPr');
	if (table.style) {
		const style = makeW(doc, 'tblStyle');
		setW(style, 'val', table.style);
		tblPr.appendChild(style);
	}
	const width = makeW(doc, 'tblW');
	setW(width, 'w', String(table.widthTwips ?? 0));
	setW(width, 'type', table.widthTwips ? 'dxa' : 'auto');
	tblPr.appendChild(width);
	if (table.alignment) {
		const jc = makeW(doc, 'jc');
		setW(jc, 'val', table.alignment);
		tblPr.appendChild(jc);
	}
	if (table.borders) tblPr.appendChild(buildBorders(doc, table.borders, 'tblBorders'));
	const look = makeW(doc, 'tblLook');
	for (const [name, value] of [
		['val', '04A0'],
		['firstRow', '1'],
		['lastRow', '0'],
		['firstColumn', '1'],
		['lastColumn', '0'],
		['noHBand', '0'],
		['noVBand', '1'],
	])
		setW(look, name, value);
	tblPr.appendChild(look);
	const tblGrid = makeW(doc, 'tblGrid');
	const columns = tableColumnCount(table);
	const widths =
		table.grid?.length === columns
			? table.grid
			: Array.from({ length: columns }, () => Math.max(1, Math.floor(contentWidthTwips / columns)));
	for (const value of widths) {
		const column = makeW(doc, 'gridCol');
		setW(column, 'w', String(value));
		tblGrid.appendChild(column);
	}
	return { tblPr, tblGrid };
}
