import { describe, expect, it } from 'vitest';
import { twips } from '@christophervr/docx-core';
import { layoutSections } from './page-flow.js';
import type { LayoutParagraph } from './input.js';
const measure = { widthOf: (text: string) => text.length * 10, lineHeightOf: () => 20 };
const page = {
	widthPx: 320,
	heightPx: 40,
	marginTopPx: 0,
	marginBottomPx: 0,
	marginLeftPx: 0,
	marginRightPx: 0,
};
const columns = {
	count: 2,
	gapPx: 20,
	widths: [
		{ widthPx: 100, gapPx: 20 },
		{ widthPx: 200, gapPx: 0 },
	],
};
const paragraph = (text: string, extra: Partial<LayoutParagraph> = {}): LayoutParagraph => ({
	kind: 'paragraph',
	id: 'p',
	runs: [{ text }],
	widowControl: false,
	...extra,
});
describe('unequal columns', () => {
	it('reflows split auto-width table cells without repeating consumed text', () => {
		const words = Array.from({ length: 24 }, (_, i) => String(i).padStart(2, '0'));
		const result = layoutSections(
			{
				sections: [
					{
						page,
						columns,
						blocks: [
							{
								kind: 'table',
								id: 't',
								rows: [{ cells: [{ paragraphs: [paragraph(words.join(' '))] }] }],
							},
						],
					},
				],
			},
			measure,
		);
		const rows = result.pages.flatMap((sheet) =>
			sheet.columns.flatMap((column) =>
				column.blocks.flatMap((box) => (box.kind === 'table' ? box.rows : [])),
			),
		);
		const lines = rows.flatMap((row) =>
			row.cells.flatMap((cell) => cell.flatMap((box) => box.lines)),
		);
		expect(
			lines
				.map((line) => line.fragments.map((fragment) => fragment.text).join(''))
				.join(' ')
				.split(/\s+/),
		).toEqual(words);
		expect(rows.slice(0, 2).map((row) => row.geometry![0]!.widthPx)).toEqual([100, 200]);
		expect(lines[2]!.fragments.filter((fragment) => fragment.text.trim())).toHaveLength(7);
	});
	it('moves a keep-together table row past earlier content and reflows it at the new width', () => {
		const result = layoutSections(
			{
				sections: [
					{
						page: { ...page, heightPx: 60 },
						columns,
						blocks: [
							paragraph('Before'),
							{
								kind: 'table',
								id: 't',
								rows: [
									{
										cantSplit: true,
										cells: [{ paragraphs: [paragraph('AB CD EF GH IJ KL MN OP QR ST')] }],
									},
								],
							},
						],
					},
				],
			},
			measure,
		);
		expect(result.pages[0]!.columns[0]!.blocks).toHaveLength(1);
		const box = result.pages[0]!.columns[1]!.blocks[0]!;
		expect(box.kind === 'table' && box.rows[0]!.geometry![0]!.widthPx).toBe(200);
		expect(box.heightPx).toBe(40);
	});
	it('reflows the remaining tokens in each width without repeating or losing text', () => {
		const words = Array.from({ length: 24 }, (_, i) => String(i).padStart(2, '0'));
		const result = layoutSections(
			{ sections: [{ page, columns, blocks: [paragraph(words.join(' '))] }] },
			measure,
		);
		const lines = result.pages.flatMap((sheet) =>
			sheet.columns.flatMap((col) =>
				col.blocks.flatMap((block) => (block.kind === 'paragraph' ? block.lines : [])),
			),
		);
		expect(
			lines
				.map((line) => line.fragments.map((fragment) => fragment.text).join(''))
				.join(' ')
				.split(/\s+/),
		).toEqual(words);
		expect(result.pages[0]!.columns.map((col) => [col.xPx, col.widthPx])).toEqual([
			[0, 100],
			[120, 200],
		]);
		expect(lines[0]!.fragments.filter((fragment) => fragment.text.trim()).length).toBe(3);
		expect(lines[2]!.fragments.filter((fragment) => fragment.text.trim()).length).toBe(7);
		expect(lines[2]!.sourceStart).toBeGreaterThan(lines[1]!.sourceEnd - 1);
	});
	it('consumes explicit column breaks once, suppresses first-line indent on continuations and retains run styles', () => {
		const p = paragraph('', {
			firstLineTwips: twips(300),
			runs: [
				{ text: 'AB', bold: true, breakAfter: 'column' },
				{ text: 'CD EF GH IJ KL', italic: true },
			],
		});
		const result = layoutSections({ sections: [{ page, columns, blocks: [p] }] }, measure);
		const narrow = result.pages[0]!.columns[0]!.blocks[0]!;
		const wide = result.pages[0]!.columns[1]!.blocks[0]!;
		expect(narrow.kind === 'paragraph' && narrow.lines[0]!.fragments[0]!.xPx).toBe(20);
		expect(wide.kind === 'paragraph' && wide.lines[0]!.fragments[0]!.xPx).toBe(0);
		expect(wide.kind === 'paragraph' && wide.lines[0]!.fragments[0]).toMatchObject({
			text: 'CD',
			italic: true,
			runIndex: 1,
		});
		expect(result.pages).toHaveLength(1);
	});
});
