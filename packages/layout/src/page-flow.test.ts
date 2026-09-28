import { describe, expect, it } from 'vitest';
import { layoutSections } from './page-flow.js';
import type {
	LayoutDocumentInput,
	LayoutPageGeometry,
	LayoutParagraph,
	LayoutTable,
} from './input.js';
import type { TextMeasurer } from './measure.js';
import { at } from './__tests__/helpers.js';

// 10px/char, 20px lines: a 100px-tall, 0-margin page holds exactly 5 lines.
const measurer: TextMeasurer = {
	widthOf: (text) => text.length * 10,
	lineHeightOf: () => 20,
};
const geometry = (widthPx: number, heightPx: number): LayoutPageGeometry => ({
	widthPx,
	heightPx,
	marginTopPx: 0,
	marginRightPx: 0,
	marginBottomPx: 0,
	marginLeftPx: 0,
});
function para(id: string, lines: number, extra: Partial<LayoutParagraph> = {}): LayoutParagraph {
	return {
		kind: 'paragraph',
		id,
		runs: [{ text: Array.from({ length: lines }, (_, i) => `L${i}`).join('\n') }],
		...extra,
	};
}
function linesOf(pageIndex: number, result: ReturnType<typeof layoutSections>) {
	return at(at(result.pages, pageIndex).columns, 0).blocks.flatMap((b) =>
		b.kind === 'paragraph' ? b.lines.length : 0,
	);
}

describe('layoutSections: page geometry per section', () => {
	it('uses each section’s own page size', () => {
		const input: LayoutDocumentInput = {
			sections: [
				{ page: geometry(200, 100), blocks: [para('p1', 1)] },
				{ page: geometry(300, 150), blocks: [para('p2', 1)] },
			],
		};
		const result = layoutSections(input, measurer);
		expect(result.pages).toHaveLength(2);
		expect(result.pages[0]).toMatchObject({ widthPx: 200, heightPx: 100 });
		expect(result.pages[1]).toMatchObject({ widthPx: 300, heightPx: 150 });
	});
});

describe('layoutSections: explicit page breaks', () => {
	it('starts a fresh page at an explicit page-break run', () => {
		const input: LayoutDocumentInput = {
			sections: [
				{
					page: geometry(200, 100),
					blocks: [
						{
							kind: 'paragraph',
							id: 'p1',
							runs: [{ text: 'before', breakAfter: 'page' }, { text: 'after' }],
						},
					],
				},
			],
		};
		const result = layoutSections(input, measurer);
		expect(result.pages).toHaveLength(2);
		expect(at(at(result.pages, 0).columns, 0).blocks).toHaveLength(1);
		expect(at(at(result.pages, 1).columns, 0).blocks).toHaveLength(1);
	});
});

describe('layoutSections: widow/orphan control', () => {
	it('pulls a line back so the next page never starts with a single stranded line (widow)', () => {
		const input: LayoutDocumentInput = {
			sections: [{ page: geometry(200, 100), blocks: [para('p1', 6)] }],
		};
		const result = layoutSections(input, measurer);
		expect(result.pages).toHaveLength(2);
		expect(linesOf(0, result)).toEqual([4]);
		expect(linesOf(1, result)).toEqual([2]);
	});

	it('never leaves a single line alone at the bottom of a page (orphan)', () => {
		const filler = para('filler', 4); // consumes 80px of the 100px page, leaving one line's worth
		const target = para('target', 3);
		const input: LayoutDocumentInput = {
			sections: [{ page: geometry(200, 100), blocks: [filler, target] }],
		};
		const result = layoutSections(input, measurer);
		expect(result.pages).toHaveLength(2);
		expect(linesOf(0, result)).toEqual([4]);
		expect(linesOf(1, result)).toEqual([3]);
	});
});

describe('layoutSections: keepNext and keepLines', () => {
	it('keeps a keepNext paragraph together with the block that follows it', () => {
		const filler = para('filler', 4); // leaves 20px (1 line) of room
		const keeper = para('keeper', 1, { keepNext: true });
		const follower = para('follower', 2);
		const input: LayoutDocumentInput = {
			sections: [{ page: geometry(200, 100), blocks: [filler, keeper, follower] }],
		};
		const result = layoutSections(input, measurer);
		// keeper alone would fit the remaining 20px, but since it must stay with
		// follower (3 lines together), the whole pair moves to the next page.
		expect(result.pages).toHaveLength(2);
		expect(linesOf(0, result)).toEqual([4]);
		expect(linesOf(1, result)).toEqual([1, 2]);
	});

	it('never splits a keepLines paragraph across pages when it fits a fresh page', () => {
		const filler = para('filler', 4); // leaves 20px of room
		const target = para('target', 3, { keepLines: true });
		const input: LayoutDocumentInput = {
			sections: [{ page: geometry(200, 100), blocks: [filler, target] }],
		};
		const result = layoutSections(input, measurer);
		expect(result.pages).toHaveLength(2);
		expect(linesOf(0, result)).toEqual([4]);
		expect(linesOf(1, result)).toEqual([3]);
	});
});

describe('layoutSections: table row fragmentation', () => {
	function table(rows: number, options: { cantSplitRow?: number } = {}): LayoutTable {
		return {
			kind: 'table',
			id: 't1',
			rows: [
				{ isHeader: true, cells: [{ paragraphs: [para('h', 1)] }] },
				...Array.from({ length: rows }, (_, i) => ({
					cells: [{ paragraphs: [para(`r${i}`, 1)] }],
					cantSplit: options.cantSplitRow === i,
				})),
			],
		};
	}

	it('repeats the header row at the top of every page the table continues onto', () => {
		// Header (20px) + 4 data rows (20px each) = 100px on page 1 exactly; the
		// remaining rows must continue onto page 2 with the header repeated.
		const input: LayoutDocumentInput = {
			sections: [{ page: geometry(200, 100), blocks: [table(6)] }],
		};
		const result = layoutSections(input, measurer);
		expect(result.pages.length).toBeGreaterThanOrEqual(2);
		const page1 = at(at(at(result.pages, 0).columns, 0).blocks, 0);
		const page2 = at(at(at(result.pages, 1).columns, 0).blocks, 0);
		expect(page1.kind).toBe('table');
		expect(page2.kind).toBe('table');
		if (page1.kind === 'table' && page2.kind === 'table') {
			expect(at(page1.rows, 0).repeated).toBe(false);
			expect(at(page2.rows, 0).repeated).toBe(true);
		}
	});

	it('moves a cantSplit row whole to the next page instead of splitting it', () => {
		const tall: LayoutTable = {
			kind: 'table',
			id: 't2',
			rows: [
				{ cells: [{ paragraphs: [para('r0', 4)] }] }, // fills the 100px page
				{ cells: [{ paragraphs: [para('r1', 2)] }], cantSplit: true },
			],
		};
		const input: LayoutDocumentInput = { sections: [{ page: geometry(200, 100), blocks: [tall] }] };
		const result = layoutSections(input, measurer);
		expect(result.pages).toHaveLength(2);
		const page2 = at(at(at(result.pages, 1).columns, 0).blocks, 0);
		expect(page2.kind).toBe('table');
		if (page2.kind === 'table') {
			expect(page2.rows).toHaveLength(1);
			expect(at(at(at(page2.rows, 0).cells, 0), 0).lines).toHaveLength(2); // whole row, not split
		}
	});
});

describe('layoutSections: odd/even section breaks and vertical alignment', () => {
	it('leaves a blank page so an odd-page section starts on an odd page', () => {
		const result = layoutSections(
			{
				sections: [
					{ page: geometry(200, 100), blocks: [para('p1', 1)] },
					{ page: geometry(200, 100), blocks: [para('p2', 1)], break: 'oddPage' },
				],
			},
			measurer,
		);
		expect(result.pages).toHaveLength(3);
		expect(at(at(result.pages, 1).columns, 0).blocks).toHaveLength(0);
		expect(result.pages[2]).toMatchObject({ index: 2, sectionIndex: 1, pageInSection: 0 });
	});

	it('does not add a blank page when the even-page section already lands on an even page', () => {
		const result = layoutSections(
			{
				sections: [
					{ page: geometry(200, 100), blocks: [para('p1', 1)] },
					{ page: geometry(200, 100), blocks: [para('p2', 1)], break: 'evenPage' },
				],
			},
			measurer,
		);
		expect(result.pages).toHaveLength(2);
	});

	it('centers and bottom-aligns content on the page', () => {
		const layout = (verticalAlign: 'center' | 'bottom') =>
			at(
				at(
					at(
						layoutSections(
							{ sections: [{ page: geometry(200, 100), blocks: [para('p1', 1)], verticalAlign }] },
							measurer,
						).pages,
						0,
					).columns,
					0,
				).blocks,
				0,
			).yPx;
		expect(layout('center')).toBe(40);
		expect(layout('bottom')).toBe(80);
	});
});

describe('layoutSections: page bottom', () => {
	it('moves a line that does not fit to the next page instead of overflowing the margin', () => {
		// The 100px page holds five 20px lines; the five-line filler leaves no room.
		const input: LayoutDocumentInput = {
			sections: [{ page: geometry(200, 100), blocks: [para('filler', 5), para('next', 1)] }],
		};
		const result = layoutSections(input, measurer);
		expect(result.pages).toHaveLength(2);
		expect(at(at(at(result.pages, 1).columns, 0).blocks, 0).blockId).toBe('next');
	});
});
