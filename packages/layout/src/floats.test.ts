import { describe, expect, it } from 'vitest';
import { layoutDocument } from './layout.js';
import { FLOAT_WRAP_NOTE } from './floats.js';
import type { LayoutDocumentInput, LayoutFloat, LayoutParagraph } from './input.js';
import type { TextMeasurer } from './measure.js';

const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };
const picture = {
	partName: 'word/media/a.png',
	contentType: 'image/png',
	widthPx: 100,
	heightPx: 50,
};

function input(blocks: LayoutParagraph[]): LayoutDocumentInput {
	return {
		sections: [
			{
				page: {
					widthPx: 800,
					heightPx: 1000,
					marginTopPx: 100,
					marginRightPx: 100,
					marginBottomPx: 100,
					marginLeftPx: 100,
				},
				blocks,
			},
		],
	};
}

const anchoredIn = (id: string, float: Partial<LayoutFloat>, lines = 1): LayoutParagraph => ({
	kind: 'paragraph',
	id,
	runs: [{ text: Array.from({ length: lines }, () => 'x').join('\n') }],
	floats: [{ ...picture, ...float }],
});

describe('pictures in pagination', () => {
	it('gives inline pictures their width on the line and raises the line to their height', () => {
		const result = layoutDocument(
			input([
				{ kind: 'paragraph', id: 'p', runs: [{ text: 'ab ' }, { text: '', object: picture }] },
			]),
			measurer,
		);
		const line = (
			result.pages[0].columns[0].blocks[0] as {
				lines: {
					heightPx: number;
					fragments: { xPx: number; widthPx: number; object?: unknown }[];
				}[];
			}
		).lines[0];
		const object = line.fragments.find((fragment) => fragment.object)!;
		expect(object).toMatchObject({ xPx: 30, widthPx: 100 });
		expect(line.heightPx).toBe(54);
	});

	it('positions floats from their reference frame, alignment and offsets', () => {
		const result = layoutDocument(
			input([
				anchoredIn('a', {
					relativeFromH: 'page',
					alignH: 'right',
					relativeFromV: 'page',
					alignV: 'bottom',
				}),
				anchoredIn('b', {
					relativeFromH: 'margin',
					alignH: 'center',
					relativeFromV: 'margin',
					offsetYPx: 30,
				}),
				anchoredIn('c', { offsetXPx: 25, offsetYPx: 5, behindText: true }),
				anchoredIn('d', {
					relativeFromH: 'leftMargin',
					alignH: 'left',
					relativeFromV: 'topMargin',
					alignV: 'center',
				}),
			]),
			measurer,
		);
		expect(result.pages[0].floats).toEqual([
			expect.objectContaining({ blockId: 'a', xPx: 700, yPx: 950 }),
			expect.objectContaining({ blockId: 'b', xPx: 350, yPx: 130 }),
			// Column-relative X and paragraph-relative Y (the third paragraph starts at 100 + 2 × 20).
			expect.objectContaining({ blockId: 'c', xPx: 125, yPx: 145, behindText: true }),
			expect.objectContaining({ blockId: 'd', xPx: 0, yPx: 25 }),
		]);
		expect(result.approximations).toContain(FLOAT_WRAP_NOTE);
	});

	it('places a float on the page where its paragraph starts', () => {
		const filler: LayoutParagraph = {
			kind: 'paragraph',
			id: 'f',
			runs: [{ text: Array(45).fill('x').join('\n') }],
		};
		const result = layoutDocument(input([filler, anchoredIn('late', { offsetYPx: 0 })]), measurer);
		expect(result.pages).toHaveLength(2);
		expect(result.pages[0].floats).toBeUndefined();
		// Five filler lines carry over, so the anchor paragraph starts at 100 + 5 × 20.
		expect(result.pages[1].floats?.[0]).toMatchObject({ blockId: 'late', yPx: 200 });
	});
});

describe('text wrapping around floats', () => {
	const words = Array.from({ length: 120 }, () => 'word').join(' ');
	const lineRights = (result: ReturnType<typeof layoutDocument>) => {
		const block = result.pages[0].columns[0].blocks[0] as {
			lines: { yPx: number; fragments: { xPx: number; widthPx: number }[] }[];
		};
		return block.lines.map((line) => ({
			y: line.yPx,
			right: Math.max(...line.fragments.map((f) => f.xPx + f.widthPx)),
			left: Math.min(...line.fragments.map((f) => f.xPx)),
		}));
	};

	it('narrows lines beside a square-wrapped picture to its wider side', () => {
		// A 100×50 picture at the right margin, 0–50px below the paragraph top.
		const paragraph = anchoredIn('p', {
			wrap: 'square',
			relativeFromH: 'margin',
			alignH: 'right',
			offsetYPx: 0,
		});
		paragraph.runs = [{ text: words }];
		const lines = lineRights(layoutDocument(input([paragraph]), measurer));
		// Beside the picture (lines at y 0 and 20, overlapping its 50px): text stops before 600 - 100 - 12.
		expect(lines[0].right).toBeLessThanOrEqual(488);
		expect(lines[2].right).toBeLessThanOrEqual(488);
		// Below it (y ≥ 50): full width again.
		expect(lines[3].right).toBeGreaterThan(488);
	});

	it('puts text on the right when the picture sits at the left', () => {
		const paragraph = anchoredIn('p', {
			wrap: 'tight',
			relativeFromH: 'margin',
			alignH: 'left',
			offsetYPx: 0,
		});
		paragraph.runs = [{ text: words }];
		const lines = lineRights(layoutDocument(input([paragraph]), measurer));
		expect(lines[0].left).toBeGreaterThanOrEqual(112);
		expect(lines[3].left).toBe(0);
	});

	it('moves text below a top-and-bottom picture, pushing later paragraphs down', () => {
		const anchor = anchoredIn('a', {
			wrap: 'topAndBottom',
			relativeFromH: 'margin',
			alignH: 'center',
			offsetYPx: 0,
		});
		const after: LayoutParagraph = { kind: 'paragraph', id: 'b', runs: [{ text: 'next' }] };
		const result = layoutDocument(input([anchor, after]), measurer);
		const [first, second] = result.pages[0].columns[0].blocks;
		// The anchor paragraph's only line starts below the 50px picture.
		expect((first as { lines: { yPx: number; gapBeforePx?: number }[] }).lines[0]).toMatchObject({
			yPx: 50,
			gapBeforePx: 50,
		});
		expect(second.yPx).toBe(70);
		expect(result.pages[0].floats?.[0]).toMatchObject({ yPx: 100 });
	});
});
