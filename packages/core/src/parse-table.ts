// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { NestedTablePreview, TableCell, TableLook } from './table-model.js';
import type { Paragraph, Table } from './model.js';
import { children, first, getW, type XmlElement, WORD_NS } from './xml.js';
import { canEditTableStructure } from './write-table.js';
import { parseShadingFill, parseShadingThemeFill, parseTableBorders } from './table-borders.js';

const dxa = (value: string | undefined): number | undefined => {
	if (value === undefined || !/^\d+$/.test(value)) return undefined;
	return Number(value);
};

function nestedPreview(node: XmlElement): NestedTablePreview {
	return {
		rows: children(node, 'tr').map((row) =>
			children(row, 'tc').map((cell) => ({
				text: Array.from(cell.getElementsByTagNameNS(WORD_NS, 't'))
					.map((t) => t.textContent ?? '')
					.join(''),
			})),
		),
	};
}

function parseCell(cell: XmlElement, parseParagraph: (p: XmlElement, id: string) => Paragraph, id: string): TableCell {
	const props = first(cell, 'tcPr');
	const result: TableCell = { paragraphs: children(cell, 'p').map((p, pi) => parseParagraph(p, `${id}p${pi}`)) };
	const gridSpan = getW(first(props, 'gridSpan'), 'val');
	if (gridSpan && /^\d+$/.test(gridSpan)) result.gridSpan = Number(gridSpan);
	const vMerge = first(props, 'vMerge');
	if (vMerge) result.verticalMerge = getW(vMerge, 'val') === 'restart' ? 'restart' : 'continue';
	const tcW = first(props, 'tcW');
	if (getW(tcW, 'type') !== 'nil') {
		const width = dxa(getW(tcW, 'w'));
		if (width !== undefined) result.widthTwips = width;
	}
	const vAlign = getW(first(props, 'vAlign'), 'val');
	if (vAlign === 'top' || vAlign === 'center' || vAlign === 'bottom') result.verticalAlign = vAlign;
	const shd = first(props, 'shd');
	const fill = parseShadingFill(shd);
	if (fill) result.shadingFill = fill;
	const themeFill = parseShadingThemeFill(shd);
	if (themeFill) result.shadingThemeFill = themeFill;
	const borders = parseTableBorders(first(props, 'tcBorders'));
	if (borders) result.borders = borders;
	const margins = first(props, 'tcMar');
	if (margins) {
		const parsed: NonNullable<TableCell['margins']> = {};
		for (const side of ['top', 'bottom', 'left', 'right'] as const) {
			const value = dxa(getW(first(margins, side), 'w'));
			if (value !== undefined) parsed[side] = value;
		}
		if (Object.keys(parsed).length) result.margins = parsed;
	}
	const nested = children(cell, 'tbl').map(nestedPreview);
	if (nested.length) result.nestedTables = nested;
	return result;
}

function attrBoolean(element: XmlElement, name: string): boolean | undefined {
	const value = getW(element, name);
	if (value === undefined) return undefined;
	return !['0', 'false', 'off', 'no', 'none'].includes(value.toLowerCase());
}

function parseLook(element: XmlElement | undefined): TableLook | undefined {
	if (!element) return undefined;
	const flag = (name: string) => attrBoolean(element, name);
	const look: TableLook = {};
	const firstRow = flag('firstRow');
	const lastRow = flag('lastRow');
	const firstColumn = flag('firstColumn');
	const lastColumn = flag('lastColumn');
	const noHBand = flag('noHBand');
	const noVBand = flag('noVBand');
	if (firstRow !== undefined) look.firstRow = firstRow;
	if (lastRow !== undefined) look.lastRow = lastRow;
	if (firstColumn !== undefined) look.firstColumn = firstColumn;
	if (lastColumn !== undefined) look.lastColumn = lastColumn;
	if (noHBand !== undefined) look.noHBand = noHBand;
	if (noVBand !== undefined) look.noVBand = noVBand;
	return Object.keys(look).length ? look : undefined;
}

/** Parses a `w:tbl`, including grid widths, merges, borders, shading and table style/look. */
export function parseTable(
	node: XmlElement,
	id: string,
	parseParagraph: (p: XmlElement, id: string) => Paragraph,
): Table {
	const rows = children(node, 'tr').map((row, ri) =>
		children(row, 'tc').map((cell, ci) => parseCell(cell, parseParagraph, `${id}-r${ri}c${ci}`)),
	);
	const table: Table = { type: 'table', id, rows, structureEditable: canEditTableStructure(node) };
	const gridElement = first(node, 'tblGrid');
	const grid = gridElement
		? children(gridElement, 'gridCol')
				.map((column) => dxa(getW(column, 'w')))
				.filter((value): value is number => value !== undefined)
		: [];
	if (grid.length) table.grid = grid;
	const tblPr = first(node, 'tblPr');
	const width = dxa(getW(first(tblPr, 'tblW'), 'w'));
	if (width !== undefined) table.widthTwips = width;
	const alignment = getW(first(tblPr, 'jc'), 'val');
	if (alignment === 'left' || alignment === 'center' || alignment === 'right') table.alignment = alignment;
	const indent = dxa(getW(first(tblPr, 'tblInd'), 'w'));
	if (indent !== undefined) table.indentTwips = indent;
	const borders = parseTableBorders(first(tblPr, 'tblBorders'));
	if (borders) table.borders = borders;
	const styleId = getW(first(tblPr, 'tblStyle'), 'val');
	if (styleId) table.style = styleId;
	const look = parseLook(first(tblPr, 'tblLook'));
	if (look) table.look = look;
	return table;
}
