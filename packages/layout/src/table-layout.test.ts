import { describe, expect, it } from 'vitest';
import { layoutRow, splitRowAtHeight } from './table-layout.js';
import type { LayoutParagraph, LayoutTableRow } from './input.js';
import type { TextMeasurer } from './measure.js';

const measurer: TextMeasurer = {
	widthOf: (text) => text.length * 10,
	lineHeightOf: () => 20,
};
const noop = () => {};
function cellParagraph(id: string, text: string): LayoutParagraph {
	return { kind: 'paragraph', id, runs: [{ text }] };
}

function threeAndTwoLineRow(): LayoutTableRow {
	return {
		cells: [
			{ paragraphs: [cellParagraph('a', 'L1\nL2\nL3')] },
			{ paragraphs: [cellParagraph('b', 'X1\nX2')] },
		],
	};
}

describe('layoutRow', () => {
	it('sizes the row to the tallest cell', () => {
		const row = layoutRow(threeAndTwoLineRow(), 200, measurer, noop);
		expect(row.heightPx).toBe(60); // 3 lines * 20px
		expect(row.cells[0][0].lines).toHaveLength(3);
		expect(row.cells[1][0].lines).toHaveLength(2);
	});

	it('splits the table width evenly across cells without explicit widths', () => {
		const row = layoutRow(
			{
				cells: [
					{ paragraphs: [cellParagraph('a', 'x')] },
					{ paragraphs: [cellParagraph('b', 'x')] },
				],
			},
			200,
			measurer,
			noop,
		);
		expect(row.cells).toHaveLength(2);
	});
});

describe('splitRowAtHeight', () => {
	it('clips every cell at the same shared line boundary, never mid-line', () => {
		const row = layoutRow(threeAndTwoLineRow(), 200, measurer, noop);
		const { before, after } = splitRowAtHeight(row, 40);
		expect(before.heightPx).toBe(40);
		expect(before.cells[0][0].lines).toHaveLength(2); // L1, L2
		expect(before.cells[1][0].lines).toHaveLength(2); // X1, X2 (fit entirely)
		expect(after).not.toBeNull();
		expect(after!.heightPx).toBe(20);
		expect(after!.cells[0][0].lines).toHaveLength(1); // L3, restarted at y=0
		expect(after!.cells[0][0].lines[0].yPx).toBe(0);
		expect(after!.cells[1]).toHaveLength(0); // nothing left over for the shorter cell
	});

	it('returns a null continuation when the whole row fits', () => {
		const row = layoutRow(threeAndTwoLineRow(), 200, measurer, noop);
		const { after } = splitRowAtHeight(row, 1000);
		expect(after).toBeNull();
	});
});
