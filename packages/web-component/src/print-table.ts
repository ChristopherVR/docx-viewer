import type {
	LayoutBorder,
	LayoutCellGeometry,
	LayoutParagraphBox,
	LayoutTableBox,
	LayoutTableRowBox,
} from 'ooxml-core/docx/layout';

const cssBorder = (border: LayoutBorder | undefined) =>
	border ? `${border.widthPx}px ${border.style} ${border.color}` : '';

/** Rows without geometry (engine-level input) share the table width evenly. */
function geometryOf(row: LayoutTableRowBox, index: number, widthPx: number): LayoutCellGeometry {
	const given = row.geometry?.[index];
	if (given) return given;
	const share = widthPx / Math.max(1, row.cells.length);
	return {
		xPx: share * index,
		widthPx: share,
		paddingLeftPx: 0,
		paddingRightPx: 0,
		contentHeightPx: 0,
	};
}

/**
 * Draws one table fragment: cells at their grid positions with shading and the document's
 * borders (no gridlines, as Word prints), content inset by cell margins and aligned vertically.
 * Each cell draws its top and left edges; the last cell of a row adds its right edge and the
 * fragment's last row its bottom, so shared edges are drawn once.
 */
export function renderTable(
	box: LayoutTableBox,
	columnWidthPx: number,
	renderParagraph: (paragraph: LayoutParagraphBox) => HTMLElement,
): HTMLElement {
	const table = document.createElement('div');
	table.className = 'dve-print-block dve-print-table';
	table.style.top = `${box.yPx}px`;
	table.style.height = `${box.heightPx}px`;
	if (box.xPx) table.style.left = `${box.xPx}px`;
	box.rows.forEach((row, rowIndex) => {
		const rowEl = document.createElement('div');
		rowEl.className = 'dve-print-row';
		rowEl.style.top = `${row.yPx}px`;
		rowEl.style.height = `${row.heightPx}px`;
		if (row.repeated) rowEl.dataset.repeated = 'true';
		const lastRow = rowIndex === box.rows.length - 1;
		row.cells.forEach((paragraphs, cellIndex) => {
			const geometry = geometryOf(row, cellIndex, columnWidthPx);
			const cell = document.createElement('div');
			cell.className = 'dve-print-cell';
			Object.assign(cell.style, {
				left: `${geometry.xPx}px`,
				width: `${geometry.widthPx}px`,
				height: `${row.heightPx}px`,
			});
			if (geometry.shading) cell.style.background = geometry.shading;
			const borders = geometry.borders;
			cell.style.borderTop = cssBorder(borders?.top);
			cell.style.borderLeft = cssBorder(borders?.left);
			if (cellIndex === row.cells.length - 1) cell.style.borderRight = cssBorder(borders?.right);
			if (lastRow) cell.style.borderBottom = cssBorder(borders?.bottom);
			const content = document.createElement('div');
			content.className = 'dve-print-cell-content';
			const free = Math.max(0, row.heightPx - geometry.contentHeightPx);
			const offset =
				geometry.verticalAlign === 'center'
					? free / 2
					: geometry.verticalAlign === 'bottom'
						? free
						: 0;
			Object.assign(content.style, {
				left: `${geometry.paddingLeftPx}px`,
				right: `${geometry.paddingRightPx}px`,
				top: `${offset}px`,
			});
			for (const paragraph of paragraphs) content.append(renderParagraph(paragraph));
			cell.append(content);
			rowEl.append(cell);
		});
		table.append(rowEl);
	});
	return table;
}
