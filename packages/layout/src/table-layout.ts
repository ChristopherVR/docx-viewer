import type { TextMeasurer } from './measure.js';
import { layoutParagraph } from './paragraph-layout.js';
import type { LayoutParagraph, LayoutTableCell, LayoutTableRow } from './input.js';
import type { LayoutCellGeometry, LayoutParagraphBox } from './result.js';

export interface RowLayout {
	heightPx: number;
	cells: LayoutParagraphBox[][];
	geometry: LayoutCellGeometry[];
}

function layoutCellParagraphs(
	paragraphs: LayoutParagraph[],
	widthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
	topPx = 0,
): { heightPx: number; boxes: LayoutParagraphBox[] } {
	let y = topPx;
	const boxes: LayoutParagraphBox[] = [];
	for (const paragraph of paragraphs) {
		const layout = layoutParagraph(paragraph, widthPx, measurer, note);
		y += layout.spacingBeforePx;
		const top = y;
		const lines = layout.lines.map((line) => ({ ...line, yPx: line.yPx + top }));
		boxes.push({
			kind: 'paragraph',
			blockId: paragraph.id,
			yPx: top,
			heightPx: layout.contentHeightPx,
			lines,
		});
		y += layout.contentHeightPx + layout.spacingAfterPx;
	}
	return { heightPx: y, boxes };
}

/**
 * Lays out one table row. Cells take their grid position and width when the table has a grid,
 * otherwise the table width split evenly, and their content is inset by the cell margins.
 */
export function layoutRow(
	row: LayoutTableRow,
	tableWidthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
): RowLayout {
	const count = row.cells.length || 1;
	let x = 0;
	const geometry: LayoutCellGeometry[] = [];
	const laidOut = row.cells.map((cell) => {
		const widthPx = cell.widthPx ?? tableWidthPx / count;
		const padding = cell.padding ?? { top: 0, right: 0, bottom: 0, left: 0 };
		const content = layoutCellParagraphs(
			cell.paragraphs,
			Math.max(1, widthPx - padding.left - padding.right),
			measurer,
			note,
			padding.top,
		);
		const heightPx = content.heightPx + padding.bottom;
		geometry.push({
			xPx: cell.xPx ?? x,
			widthPx,
			paddingLeftPx: padding.left,
			paddingRightPx: padding.right,
			contentHeightPx: heightPx,
			...(cell.borders ? { borders: cell.borders } : {}),
			...(cell.shading ? { shading: cell.shading } : {}),
			...(cell.verticalAlign ? { verticalAlign: cell.verticalAlign } : {}),
		});
		x = (cell.xPx ?? x) + widthPx;
		return { heightPx, boxes: content.boxes };
	});
	return {
		heightPx: Math.max(0, ...laidOut.map((cell) => cell.heightPx)),
		cells: laidOut.map((cell) => cell.boxes),
		geometry,
	};
}

interface CellSplit {
	before: LayoutParagraphBox[];
	after: LayoutParagraphBox[];
	beforeHeightPx: number;
}

/** Splits one cell's already-laid-out paragraph boxes at `cutHeightPx`, never mid-line. */
function splitCellAtHeight(boxes: LayoutParagraphBox[], cutHeightPx: number): CellSplit {
	const before: LayoutParagraphBox[] = [];
	const after: LayoutParagraphBox[] = [];
	let afterOrigin: number | null = null;
	let beforeHeightPx = 0;
	for (const box of boxes) {
		if (afterOrigin !== null) {
			after.push({ ...box, yPx: box.yPx - afterOrigin });
			continue;
		}
		if (box.yPx + box.heightPx <= cutHeightPx) {
			before.push(box);
			beforeHeightPx = box.yPx + box.heightPx;
			continue;
		}
		if (box.yPx >= cutHeightPx) {
			afterOrigin = box.yPx;
			after.push({ ...box, yPx: 0 });
			continue;
		}
		const beforeLines = box.lines.filter(
			(line) => box.yPx + line.yPx + line.heightPx <= cutHeightPx,
		);
		const afterLines = box.lines.slice(beforeLines.length);
		if (beforeLines.length) {
			const bottom = beforeLines.at(-1)!.yPx + beforeLines.at(-1)!.heightPx;
			before.push({ ...box, heightPx: bottom, lines: beforeLines });
			beforeHeightPx = box.yPx + bottom;
		}
		if (afterLines.length) {
			const shift = afterLines[0].yPx;
			afterOrigin = box.yPx + shift;
			after.push({
				...box,
				yPx: 0,
				heightPx: afterLines.at(-1)!.yPx + afterLines.at(-1)!.heightPx - shift,
				lines: afterLines.map((line) => ({ ...line, yPx: line.yPx - shift })),
			});
		} else {
			afterOrigin = box.yPx + box.heightPx;
		}
	}
	return { before, after, beforeHeightPx };
}

export interface RowSplit {
	before: RowLayout;
	after: RowLayout | null;
}

/**
 * Splits a row that does not fit in `cutHeightPx` of remaining page/column
 * space. Every cell is clipped at the same shared line boundary (never
 * mid-line); a cell shorter than the cut simply has nothing in `after`.
 * Returns `after: null` when every cell fit (nothing left to continue).
 */
export function splitRowAtHeight(row: RowLayout, cutHeightPx: number): RowSplit {
	const splits = row.cells.map((cell) => splitCellAtHeight(cell, cutHeightPx));
	const anyAfter = splits.some((split) => split.after.length > 0);
	const before: RowLayout = {
		heightPx: Math.max(0, ...splits.map((split) => split.beforeHeightPx)),
		cells: splits.map((split) => split.before),
		geometry: row.geometry,
	};
	if (!anyAfter) return { before, after: null };
	const after: RowLayout = {
		heightPx: Math.max(
			0,
			...splits.map((split) =>
				split.after.reduce((max, box) => Math.max(max, box.yPx + box.heightPx), 0),
			),
		),
		cells: splits.map((split) => split.after),
		geometry: row.geometry,
	};
	return { before, after };
}
