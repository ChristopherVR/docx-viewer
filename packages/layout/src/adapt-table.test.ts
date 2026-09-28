import { describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel, type Table } from '@christophervr/docx-core';
import { adaptDocumentModel } from './adapter.js';
import { layoutDocument } from './layout.js';
import type { LayoutTable } from './input.js';
import type { LayoutTableBox } from './result.js';
import type { TextMeasurer } from './measure.js';

const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };
const cell = (text: string, extra: Partial<Table['rows'][number][number]> = {}) => ({
	paragraphs: [{ type: 'paragraph' as const, id: `c-${text}`, runs: [{ text }] }],
	...extra,
});

function model(table: Table): DocumentModel {
	return { ...createDocument(), blocks: [table] };
}

const single = { style: 'single', sizeEighthPoints: 8, color: 'FF0000' };
const table: Table = {
	type: 'table',
	id: 't',
	grid: [1500, 1500, 3000],
	alignment: 'center',
	borders: {
		top: single,
		bottom: single,
		left: single,
		right: single,
		insideH: single,
		insideV: single,
	},
	rows: [
		[cell('a', { gridSpan: 2, verticalMerge: 'restart' }), cell('b', { shadingFill: 'FFFF00' })],
		[
			cell('', { gridSpan: 2, verticalMerge: 'continue' }),
			cell('c', { verticalAlign: 'bottom', margins: { left: 300 } }),
		],
	],
};

describe('tables in Print Layout', () => {
	it('adapts grid positions, spans, margins, borders, shading and vertical merges', () => {
		const adapted = adaptDocumentModel(model(table)).sections[0].blocks[0] as LayoutTable;
		expect(adapted).toMatchObject({ widthPx: 400, alignment: 'center' });
		const [first, second] = adapted.rows;
		expect(first.cells[0]).toMatchObject({
			xPx: 0,
			widthPx: 200,
			padding: { left: 7.2, right: 7.2 },
		});
		expect(first.cells[1]).toMatchObject({ xPx: 200, widthPx: 200, shading: '#FFFF00' });
		expect(first.cells[0].borders?.top).toEqual({
			widthPx: 4 / 3,
			style: 'solid',
			color: '#FF0000',
		});
		// The merged cells hide the border between them.
		expect(first.cells[0].borders?.bottom).toBeUndefined();
		expect(second.cells[0].borders?.top).toBeUndefined();
		expect(second.cells[1]).toMatchObject({ verticalAlign: 'bottom', padding: { left: 20 } });
	});

	it('centers the table in its column and positions cells on the grid', () => {
		const result = layoutDocument(adaptDocumentModel(model(table)), measurer);
		const box = result.pages[0].columns[0].blocks[0] as LayoutTableBox;
		// The 624px content column (816 − 2 × 96) centers the 400px table.
		expect(box.xPx).toBe(112);
		expect(box.rows[0].geometry?.map((g) => [g.xPx, g.widthPx])).toEqual([
			[0, 200],
			[200, 200],
		]);
		const text = box.rows[1].cells[1][0];
		expect(text.lines[0].fragments[0].xPx).toBe(0);
		expect(box.rows[1].geometry?.[1]).toMatchObject({ paddingLeftPx: 20, verticalAlign: 'bottom' });
	});
});

describe('table rows in Print Layout', () => {
	it('applies row heights, table cell margins and repeats header rows on the next page', () => {
		const rows = Array.from({ length: 40 }, (_, index) => [cell(`r${index}`)]);
		const tall: Table = {
			type: 'table',
			id: 'tall',
			grid: [3000],
			cellMargins: { left: 0, right: 0 },
			rowProperties: [
				{ header: true, heightTwips: 600, heightRule: 'exact' },
				{ heightTwips: 900 },
			],
			rows,
		};
		const result = layoutDocument(adaptDocumentModel(model(tall)), measurer);
		const first = result.pages[0].columns[0].blocks[0] as LayoutTableBox;
		// Exact 600 twips = 40px; at-least 900 twips = 60px (more than one 20px line).
		expect(first.rows.slice(0, 2).map((row) => row.heightPx)).toEqual([40, 60]);
		expect(first.rows[0].geometry?.[0].paddingLeftPx).toBe(0);
		const next = result.pages[1].columns[0].blocks[0] as LayoutTableBox;
		expect(next.rows[0]).toMatchObject({ repeated: true, heightPx: 40 });
	});
});
