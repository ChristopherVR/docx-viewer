// @vitest-environment jsdom
import type { LayoutResult } from '@christophervr/docx-layout';
import { describe, expect, it } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import { renderPrintLayout } from './print-layout';
import { setLineNumbering } from './section-commands';

const line = (yPx: number) => ({
	yPx,
	heightPx: 20,
	sourceStart: 0,
	sourceEnd: 1,
	fragments: [{ text: 'x', xPx: 0, widthPx: 8, runIndex: 0 }],
});
const page = (index: number, sectionIndex: number): LayoutResult['pages'][number] => ({
	index,
	sectionIndex,
	pageInSection: index,
	widthPx: 300,
	heightPx: 200,
	marginTopPx: 10,
	marginRightPx: 10,
	marginBottomPx: 10,
	marginLeftPx: 60,
	columns: [
		{
			xPx: 0,
			widthPx: 200,
			blocks: [
				{
					kind: 'paragraph',
					blockId: `p${index}`,
					yPx: 0,
					heightPx: 60,
					lines: [line(0), line(20), line(40)],
				},
			],
		},
	],
});
const result: LayoutResult = { approximations: [], pages: [page(0, 0), page(1, 0)] };
const numbers = (root: HTMLElement) =>
	[...root.querySelectorAll('.dve-print-line-number')].map((el) => el.textContent);

describe('line numbers in Print Layout', () => {
	it('counts on across pages when continuous and restarts each page when asked', () => {
		const continuous = renderPrintLayout(result, undefined, {
			lineNumbers: [{ countBy: 1, start: 1, restart: 'continuous' }],
		});
		expect(numbers(continuous.element)).toEqual(['1', '2', '3', '4', '5', '6']);
		const perPage = renderPrintLayout(result, undefined, {
			lineNumbers: [{ countBy: 1, start: 1, restart: 'newPage' }],
		});
		expect(numbers(perPage.element)).toEqual(['1', '2', '3', '1', '2', '3']);
	});

	it('shows every countBy-th line and draws nothing without settings', () => {
		const every = renderPrintLayout(result, undefined, {
			lineNumbers: [{ countBy: 2, start: 1, restart: 'continuous' }],
		});
		expect(numbers(every.element)).toEqual(['2', '4', '6']);
		expect(numbers(renderPrintLayout(result).element)).toEqual([]);
	});
});

describe('setLineNumbering', () => {
	it('turns line numbers on with a restart rule and off again', () => {
		const model = createDocument();
		const on = setLineNumbering(model, 0, 'newSection');
		expect(on.sections?.[0]).toMatchObject({
			lineNumbering: true,
			lineNumberSettings: { countBy: 1, start: 1, restart: 'newSection' },
		});
		const off = setLineNumbering(on, 0, 'none');
		expect(off.sections?.[0]?.lineNumbering).toBeUndefined();
		expect(off.sections?.[0]?.lineNumberSettings).toBeUndefined();
	});
});
