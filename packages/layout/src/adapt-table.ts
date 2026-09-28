import {
	resolveCellVisuals,
	resolveThemeColorReference,
	type DocumentModel,
	type Paragraph,
	type Table,
	type TableBorderSide,
} from '@christophervr/docx-core';
import type {
	LayoutBorder,
	LayoutCellBorders,
	LayoutCellPadding,
	LayoutParagraph,
	LayoutTable,
	LayoutTableCell,
} from './input.js';

const twipsToPx = (twips: number): number => twips / 15;
/** Word's default table cell margins (`w:tblCellMar`): 108 twips (0.075in) left and right. */
const DEFAULT_PADDING_TWIPS = { top: 0, right: 108, bottom: 0, left: 108 };

function border(
	side: TableBorderSide | undefined,
	theme: DocumentModel['theme'],
): LayoutBorder | undefined {
	if (!side?.style || side.style === 'none' || side.style === 'nil') return undefined;
	const style: LayoutBorder['style'] = side.style.startsWith('double')
		? 'double'
		: /dot/i.test(side.style)
			? 'dotted'
			: /dash/i.test(side.style)
				? 'dashed'
				: 'solid';
	const themed =
		side.themeColor && theme
			? resolveThemeColorReference({ token: side.themeColor }, theme)
			: undefined;
	const hex = side.color && /^[0-9a-f]{6}$/i.test(side.color) ? `#${side.color}` : undefined;
	// `sz` is in eighths of a point; a double line needs room for both strokes.
	const widthPx = Math.max(
		style === 'double' ? 3 : 1,
		(((side.sizeEighthPoints ?? 4) / 8) * 96) / 72,
	);
	return { widthPx, style, color: themed ?? hex ?? '#000000' };
}

/**
 * Converts a model table for pagination: grid positions and widths (merged cells span grid
 * columns), cell margins, resolved borders and shading (table style, table and cell layers), and
 * vertical merges drawn as one cell by hiding the borders between the merged cells.
 */
export function adaptTable(
	table: Table,
	model: DocumentModel,
	adaptParagraph: (paragraph: Paragraph) => LayoutParagraph,
): LayoutTable {
	const grid = table.grid?.map(twipsToPx);
	const columnCount = grid?.length ?? Math.max(1, ...table.rows.map((row) => row.length));
	const gridWidth = (start: number, span: number) =>
		grid ? grid.slice(start, start + span).reduce((sum, width) => sum + width, 0) : undefined;
	// Grid column of each cell, so vertical merges line up across rows with different spans.
	const gridIndexes = table.rows.map((row) => {
		let index = 0;
		return row.map((cell) => {
			const start = index;
			index += cell.gridSpan ?? 1;
			return start;
		});
	});
	const cellAt = (rowIndex: number, gridIndex: number) => {
		const position = gridIndexes[rowIndex]?.indexOf(gridIndex) ?? -1;
		return position >= 0 ? table.rows[rowIndex][position] : undefined;
	};
	return {
		kind: 'table',
		id: table.id,
		...(grid ? { widthPx: grid.reduce((sum, width) => sum + width, 0) } : {}),
		...(table.indentTwips ? { indentPx: twipsToPx(table.indentTwips) } : {}),
		...(table.alignment ? { alignment: table.alignment } : {}),
		rows: table.rows.map((row, rowIndex) => ({
			cells: row.map((cell, cellIndex): LayoutTableCell => {
				const start = gridIndexes[rowIndex][cellIndex];
				const span = cell.gridSpan ?? 1;
				const visuals = resolveCellVisuals(
					table,
					cell,
					{
						row: rowIndex,
						lastRow: rowIndex,
						column: start,
						lastColumn: start + span - 1,
						rowCount: table.rows.length,
						columnCount,
					},
					model.tableStyles,
				);
				const borders: LayoutCellBorders = {};
				for (const side of ['top', 'right', 'bottom', 'left'] as const) {
					const resolved = border(visuals.borders?.[side], model.theme);
					if (resolved) borders[side] = resolved;
				}
				if (cell.verticalMerge === 'continue') delete borders.top;
				if (cellAt(rowIndex + 1, start)?.verticalMerge === 'continue') delete borders.bottom;
				const margins = { ...DEFAULT_PADDING_TWIPS, ...cell.margins };
				const padding: LayoutCellPadding = {
					top: twipsToPx(margins.top),
					right: twipsToPx(margins.right),
					bottom: twipsToPx(margins.bottom),
					left: twipsToPx(margins.left),
				};
				const width =
					gridWidth(start, span) ?? (cell.widthTwips ? twipsToPx(cell.widthTwips) : undefined);
				const fill = visuals.shadingFill;
				return {
					paragraphs: cell.paragraphs.map(adaptParagraph),
					...(width !== undefined ? { widthPx: width } : {}),
					...(grid ? { xPx: gridWidth(0, start) } : {}),
					padding,
					...(Object.keys(borders).length ? { borders } : {}),
					...(fill && /^[0-9a-f]{6}$/i.test(fill) ? { shading: `#${fill}` } : {}),
					...(cell.verticalAlign ? { verticalAlign: cell.verticalAlign } : {}),
				};
			}),
		})),
	};
}
