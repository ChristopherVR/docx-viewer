import {
	resolveTableStyleFormatting,
	type Table,
	type TableBorderSide,
	type TableBorders,
	type TableStyleCatalog,
} from '@christophervr/docx-core';

type Side = 'top' | 'bottom' | 'left' | 'right';
const NONE: TableBorderSide = { style: 'none' };

/** The table-level border a cell edge shows: the outer edge on the boundary, inside lines otherwise. */
function edgeFrom(borders: TableBorders | undefined, side: Side, onBoundary: boolean) {
	if (!borders) return undefined;
	if (onBoundary) return borders[side];
	return side === 'top' || side === 'bottom' ? borders.insideH : borders.insideV;
}

export interface CellPlacement {
	row: number;
	lastRow: number;
	column: number;
	lastColumn: number;
	rowCount: number;
	columnCount: number;
}

/**
 * Display-only borders and shading for one cell, layered like Word: table style (with its
 * conditional regions), then the table's direct borders, then the cell's own `tcPr`. Returns
 * undefined borders when the table carries no border information at all, so gridlines show.
 */
export function resolveCellVisuals(
	table: Table,
	cell: Table['rows'][number][number],
	placement: CellPlacement,
	styles: TableStyleCatalog | undefined,
): { borders?: Record<Side, TableBorderSide>; shadingFill?: string } {
	const styled = table.style
		? resolveTableStyleFormatting(
				table.style,
				styles,
				table.look,
				placement.row,
				placement.rowCount,
				placement.column,
				placement.columnCount,
			)
		: {};
	const hasBorderInfo = Boolean(styled.borders || table.borders || cell.borders);
	const shadingFill = cell.shadingFill ?? styled.shadingFill;
	if (!hasBorderInfo) return shadingFill ? { shadingFill } : {};
	const boundary: Record<Side, boolean> = {
		top: placement.row === 0,
		bottom: placement.lastRow === placement.rowCount - 1,
		left: placement.column === 0,
		right: placement.lastColumn === placement.columnCount - 1,
	};
	const borders = {} as Record<Side, TableBorderSide>;
	for (const side of ['top', 'bottom', 'left', 'right'] as const)
		borders[side] =
			cell.borders?.[side] ??
			edgeFrom(table.borders, side, boundary[side]) ??
			edgeFrom(styled.borders, side, boundary[side]) ??
			NONE;
	return { borders, ...(shadingFill ? { shadingFill } : {}) };
}
