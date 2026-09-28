import type { PageCursor } from './page-cursor.js';
import { layoutRow, splitRowAtHeight, type RowLayout } from './table-layout.js';
import type { TextMeasurer } from './measure.js';
import type { LayoutTable } from './input.js';
import type { LayoutTableBox, LayoutTableRowBox } from './result.js';

/**
 * Places a table row by row, splitting a row across a page/column boundary
 * unless it is `cantSplit`, and repeating `tblHeader` rows at the top of
 * every page/column fragment the table continues onto. Each fragment the
 * table touches becomes its own `LayoutTableBox` sharing the table's
 * `blockId`, since a single result box can only live in one column.
 */
export function placeTable(
	cursor: PageCursor,
	table: LayoutTable,
	measurer: TextMeasurer,
	note: (message: string) => void,
): void {
	const headerRows = table.rows.filter((row) => row.isHeader);
	const headerLayouts = headerRows.map((row) =>
		layoutRow(row, cursor.columnWidthPx, measurer, note),
	);

	// Table indent, or alignment of the grid within the column.
	const width = table.widthPx ?? cursor.columnWidthPx;
	const xPx =
		table.alignment === 'center'
			? Math.max(0, (cursor.columnWidthPx - width) / 2)
			: table.alignment === 'right'
				? Math.max(0, cursor.columnWidthPx - width)
				: (table.indentPx ?? 0);
	let fragmentRows: LayoutTableRowBox[] = [];
	let fragmentHeight = 0;

	const toRowBox = (layout: RowLayout, repeated: boolean): LayoutTableRowBox => {
		const box: LayoutTableRowBox = {
			yPx: fragmentHeight,
			heightPx: layout.heightPx,
			repeated,
			cells: layout.cells,
			geometry: layout.geometry,
		};
		fragmentHeight += layout.heightPx;
		return box;
	};

	function flushFragment() {
		if (!fragmentRows.length) return;
		const box: LayoutTableBox = {
			kind: 'table',
			blockId: table.id,
			...(xPx ? { xPx } : {}),
			yPx: 0,
			heightPx: fragmentHeight,
			rows: fragmentRows,
		};
		cursor.place(box, fragmentHeight);
		fragmentRows = [];
		fragmentHeight = 0;
	}
	function startNewFragment() {
		flushFragment();
		cursor.newColumn();
		for (const header of headerLayouts) fragmentRows.push(toRowBox(header, true));
	}

	function placeRow(initial: RowLayout, cantSplit: boolean) {
		let remainingRow: RowLayout | null = initial;
		while (remainingRow) {
			const remaining = cursor.remainingHeightPx() - fragmentHeight;
			if (remainingRow.heightPx <= remaining) {
				fragmentRows.push(toRowBox(remainingRow, false));
				remainingRow = null;
				continue;
			}
			if (cantSplit) {
				if (fragmentRows.length) startNewFragment();
				const freshRemaining = cursor.remainingHeightPx() - fragmentHeight;
				if (remainingRow.heightPx > freshRemaining)
					note('A table row marked "keep row together" is taller than an empty page/column.');
				fragmentRows.push(toRowBox(remainingRow, false));
				remainingRow = null;
				continue;
			}
			const { before, after } = splitRowAtHeight(remainingRow, Math.max(0, remaining));
			if (before.cells.some((cell) => cell.length)) fragmentRows.push(toRowBox(before, false));
			startNewFragment();
			remainingRow = after;
		}
	}

	for (const row of table.rows) {
		const layout = layoutRow(row, cursor.columnWidthPx, measurer, note);
		placeRow(layout, Boolean(row.cantSplit) || row.isHeader === true);
	}
	flushFragment();
}
