// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import type { LayoutResult } from '@christophervr/ooxml-core/docx/layout';
import { renderPrintLayout } from './print-layout';
import { at } from './test-support';

function sampleResult(): LayoutResult {
	return {
		approximations: ['example approximation'],
		pages: [
			{
				index: 0,
				sectionIndex: 0,
				pageInSection: 0,
				widthPx: 300,
				heightPx: 200,
				marginTopPx: 10,
				marginRightPx: 10,
				marginBottomPx: 10,
				marginLeftPx: 10,
				columns: [
					{
						xPx: 0,
						widthPx: 280,
						blocks: [
							{
								kind: 'paragraph',
								blockId: 'p1',
								yPx: 0,
								heightPx: 20,
								lines: [
									{
										yPx: 0,
										heightPx: 20,
										sourceStart: 0,
										sourceEnd: 5,
										fragments: [{ text: 'hello', xPx: 0, widthPx: 40, runIndex: 0 }],
									},
								],
							},
						],
					},
				],
			},
		],
	};
}

describe('renderPrintLayout', () => {
	it('renders scaled text with independent spacing and threshold-controlled kerning', () => {
		const result = sampleResult();
		const block = result.pages[0]!.columns[0]!.blocks[0]!;
		if (block.kind !== 'paragraph') throw new Error('Expected paragraph');
		Object.assign(block.lines[0]!.fragments[0]!, {
			fontSizePt: 12,
			textScalePercent: 200,
			characterSpacingPx: 2,
			kerningThresholdPt: 12,
			topPx: 4,
			boxHeightPx: 20,
		});
		const rendered = renderPrintLayout(result);
		const span = rendered.element.querySelector('.dve-print-line span') as HTMLElement;
		expect(span.style.transform).toBe('scaleX(2)');
		expect(span.style.transformOrigin).toBe('left center');
		expect(span.style.letterSpacing).toBe('1px');
		expect(span.style.fontKerning).toBe('normal');
		expect(span.style.top).toBe('4px');
		expect(span.style.lineHeight).toBe('20px');
	});
	it('draws a column separator halfway across the gap', () => {
		const result = sampleResult();
		const page = result.pages[0]!;
		page.columns = [
			{ xPx: 0, widthPx: 120, blocks: [] },
			{ xPx: 160, widthPx: 120, blocks: [] },
		];
		page.columnSeparator = true;
		const rendered = renderPrintLayout(result);
		const rule = rendered.element.querySelector<HTMLElement>('.dve-print-column-rule')!;
		expect(rule.style.left).toBe('150px');
		expect(rule.style.height).toBe('180px');
	});
	it('renders one sheet per page sized from the page geometry', () => {
		const handle = renderPrintLayout(sampleResult());
		const pages = handle.element.querySelectorAll('.dve-print-page');
		expect(pages).toHaveLength(1);
		expect(handle.pageCount).toBe(1);
		expect((pages[0] as HTMLElement).style.width).toBe('300px');
		expect((pages[0] as HTMLElement).style.height).toBe('200px');
	});

	it('renders paragraph lines with their fragment text', () => {
		const handle = renderPrintLayout(sampleResult());
		const line = handle.element.querySelector('.dve-print-line')!;
		expect(line.textContent).toBe('hello');
	});

	it('maps a click on a rendered line back to a blockId/offset', () => {
		const handle = renderPrintLayout(sampleResult());
		const line = handle.element.querySelector('.dve-print-line') as HTMLElement;
		document.body.append(handle.element);
		const hit = handle.resolveClick(line, 0);
		expect(hit).not.toBeNull();
		expect(hit?.blockId).toBe('p1');
		expect(hit?.offset).toBeGreaterThanOrEqual(0);
		expect(hit?.offset).toBeLessThanOrEqual(5);
		handle.element.remove();
	});

	it('returns null when the clicked element is not part of a rendered line', () => {
		const handle = renderPrintLayout(sampleResult());
		expect(handle.resolveClick(document.body, 0)).toBeNull();
	});
});

describe('Print Layout pictures', () => {
	const picture = {
		partName: 'word/media/a.png',
		contentType: 'image/png',
		widthPx: 60,
		heightPx: 30,
	};

	it('draws inline pictures on their line and floating pictures at their page position', () => {
		const result = sampleResult();
		const page = at(result.pages, 0);
		const paragraph = at(at(page.columns, 0).blocks, 0);
		if (paragraph.kind === 'paragraph')
			at(paragraph.lines, 0).fragments.push({
				text: '',
				xPx: 40,
				widthPx: 60,
				runIndex: 1,
				object: picture,
			});
		page.floats = [
			{ blockId: 'p1', xPx: 200, yPx: 150, ...picture, behindText: false },
			{
				blockId: 'p1',
				xPx: 5,
				yPx: 5,
				...picture,
				partName: 'word/media/missing.png',
				behindText: true,
			},
		];
		const urls: Record<string, string> = { 'word/media/a.png': 'blob:a' };
		const { element } = renderPrintLayout(result, (partName) => urls[partName]);
		const inline = element.querySelector<HTMLImageElement>(
			'.dve-print-line img.dve-print-picture',
		)!;
		expect(inline.getAttribute('src')).toBe('blob:a');
		expect([inline.style.left, inline.style.width, inline.style.height]).toEqual([
			'40px',
			'60px',
			'30px',
		]);
		const floats = element.querySelectorAll<HTMLElement>('.dve-print-page > .dve-print-float');
		expect(floats).toHaveLength(2);
		const [firstFloat, secondFloat] = [at(floats, 0), at(floats, 1)];
		expect([firstFloat.style.left, firstFloat.style.top]).toEqual(['200px', '150px']);
		expect(secondFloat.classList.contains('dve-print-float-behind')).toBe(true);
		expect(secondFloat.classList.contains('dve-print-picture-missing')).toBe(true);
	});
});

describe('Print Layout tab leaders', () => {
	it('fills a tab with its leader across the tab width', () => {
		const result = sampleResult();
		const paragraph = at(at(at(result.pages, 0).columns, 0).blocks, 0);
		if (paragraph.kind === 'paragraph')
			at(paragraph.lines, 0).fragments.push(
				{ text: '', xPx: 40, widthPx: 120, runIndex: 0, leader: 'dot' },
				{ text: '', xPx: 160, widthPx: 30, runIndex: 0, leader: 'underscore' },
			);
		const { element } = renderPrintLayout(result);
		const leaders = element.querySelectorAll<HTMLElement>('.dve-print-leader');
		const [dots, rule] = [at(leaders, 0), at(leaders, 1)];
		expect(dots.style.width).toBe('120px');
		expect(dots.textContent).toMatch(/^\.{10,}$/);
		expect(rule.textContent).toBe('');
		expect(rule.style.borderBottom).toContain('solid');
	});
});

describe('Print Layout tables', () => {
	it('draws cells at their grid positions with borders, shading and vertical alignment', () => {
		const line = {
			yPx: 0,
			heightPx: 20,
			sourceStart: 0,
			sourceEnd: 1,
			fragments: [{ text: 'x', xPx: 0, widthPx: 10, runIndex: 0 }],
		};
		const paragraph = {
			kind: 'paragraph' as const,
			blockId: 'c',
			yPx: 0,
			heightPx: 20,
			lines: [line],
		};
		const red = { widthPx: 1, style: 'solid' as const, color: '#ff0000' };
		const result = sampleResult();
		at(at(result.pages, 0).columns, 0).blocks = [
			{
				kind: 'table',
				blockId: 't',
				xPx: 40,
				yPx: 0,
				heightPx: 60,
				rows: [
					{
						yPx: 0,
						heightPx: 60,
						repeated: false,
						cells: [[paragraph], [paragraph]],
						geometry: [
							{
								xPx: 0,
								widthPx: 100,
								paddingLeftPx: 7,
								paddingRightPx: 7,
								contentHeightPx: 20,
								borders: { top: red, left: red, right: red, bottom: red },
							},
							{
								xPx: 100,
								widthPx: 150,
								paddingLeftPx: 7,
								paddingRightPx: 7,
								contentHeightPx: 20,
								shading: '#ffff00',
								verticalAlign: 'bottom',
								borders: { top: red, left: red, right: red, bottom: red },
							},
						],
					},
				],
			},
		];
		const { element } = renderPrintLayout(result);
		const table = element.querySelector<HTMLElement>('.dve-print-table')!;
		expect(table.style.left).toBe('40px');
		const cells = element.querySelectorAll<HTMLElement>('.dve-print-cell');
		const [first, second] = [at(cells, 0), at(cells, 1)];
		expect([first.style.left, first.style.width, second.style.left]).toEqual([
			'0px',
			'100px',
			'100px',
		]);
		// Shared edges are drawn once: the first cell has no right edge, the last does.
		expect(first.style.borderRight).toBe('');
		expect(second.style.borderRight).toBe('1px solid rgb(255, 0, 0)');
		expect(second.style.background).toContain('rgb(255, 255, 0)');
		const content = second.querySelector<HTMLElement>('.dve-print-cell-content')!;
		expect([content.style.left, content.style.top]).toEqual(['7px', '40px']);
	});
});

describe('Print Layout footnotes', () => {
	it('draws the footnote area above the bottom margin with a separator rule', () => {
		const result = sampleResult();
		const line = {
			yPx: 0,
			heightPx: 20,
			sourceStart: 0,
			sourceEnd: 1,
			fragments: [{ text: '1', xPx: 0, widthPx: 6, runIndex: 0, script: 'super' as const }],
		};
		at(result.pages, 0).footnotes = [
			{
				id: '1',
				yPx: 0,
				heightPx: 20,
				paragraphs: [{ kind: 'paragraph', blockId: 'n1', yPx: 0, heightPx: 20, lines: [line] }],
			},
		];
		const { element } = renderPrintLayout(result);
		const area = element.querySelector<HTMLElement>('.dve-print-footnotes')!;
		// 200px page − 10px bottom margin − 20px of notes.
		expect([area.style.top, area.style.left]).toEqual(['170px', '10px']);
		expect(area.querySelector('.dve-print-footnote-separator')).not.toBeNull();
		const mark = area.querySelector<HTMLElement>('.dve-print-line span')!;
		expect(mark.textContent).toBe('1');
		expect(mark.style.fontSize).toBe('7.15pt');
	});
});
