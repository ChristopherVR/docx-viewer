// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import type { LayoutResult } from '@christophervr/docx-layout';
import { renderPrintLayout } from './print-layout';

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
		const page = result.pages[0];
		const paragraph = page.columns[0].blocks[0];
		if (paragraph.kind === 'paragraph')
			paragraph.lines[0].fragments.push({
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
		expect([floats[0].style.left, floats[0].style.top]).toEqual(['200px', '150px']);
		expect(floats[1].classList.contains('dve-print-float-behind')).toBe(true);
		expect(floats[1].classList.contains('dve-print-picture-missing')).toBe(true);
	});
});

describe('Print Layout tab leaders', () => {
	it('fills a tab with its leader across the tab width', () => {
		const result = sampleResult();
		const paragraph = result.pages[0].columns[0].blocks[0];
		if (paragraph.kind === 'paragraph')
			paragraph.lines[0].fragments.push(
				{ text: '', xPx: 40, widthPx: 120, runIndex: 0, leader: 'dot' },
				{ text: '', xPx: 160, widthPx: 30, runIndex: 0, leader: 'underscore' },
			);
		const { element } = renderPrintLayout(result);
		const [dots, rule] = element.querySelectorAll<HTMLElement>('.dve-print-leader');
		expect(dots.style.width).toBe('120px');
		expect(dots.textContent).toMatch(/^\.{10,}$/);
		expect(rule.textContent).toBe('');
		expect(rule.style.borderBottom).toContain('solid');
	});
});
