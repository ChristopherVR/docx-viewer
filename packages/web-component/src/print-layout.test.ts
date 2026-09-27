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
