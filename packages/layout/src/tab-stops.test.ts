import { describe, expect, it } from 'vitest';
import type { DocumentModel } from '@christophervr/docx-core';
import { adaptDocumentModel } from './adapter.js';
import type { LayoutParagraph } from './input.js';
import type { TextMeasurer } from './measure.js';
import { layoutParagraph } from './paragraph-layout.js';
import { placeTab } from './tab-stops.js';

// 10px per character, 20px lines.
const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };
const noop = () => {};
const paragraph = (text: string, extra: Partial<LayoutParagraph> = {}): LayoutParagraph => ({
	kind: 'paragraph',
	id: 'p',
	runs: [{ text }],
	...extra,
});
const fragmentsOf = (result: ReturnType<typeof layoutParagraph>) =>
	result.lines[0].fragments.map(({ text, xPx, widthPx, leader }) => ({
		text,
		xPx,
		widthPx,
		...(leader ? { leader } : {}),
	}));

describe('tab stops', () => {
	it('advances to custom, hanging and default stops', () => {
		const stops = [{ posPx: 100, align: 'left' as const }];
		expect(placeTab(30, stops, undefined, 0, 0)).toEqual({ widthPx: 70 });
		expect(placeTab(30, [], 48, 0, 0)).toEqual({ widthPx: 18 });
		// Past the last custom stop, default stops every half inch (48px) apply.
		expect(placeTab(110, stops, undefined, 0, 0)).toEqual({ widthPx: 34 });
		expect(placeTab(10, [{ posPx: 100, align: 'clear' }], undefined, 0, 0)).toEqual({
			widthPx: 38,
		});
	});

	it('right-aligns text at a right stop with a dot leader, like a TOC entry', () => {
		const result = layoutParagraph(
			paragraph('Intro\t12', { tabStops: [{ posPx: 200, align: 'right', leader: 'dot' }] }),
			200,
			measurer,
			noop,
		);
		expect(fragmentsOf(result)).toEqual([
			{ text: 'Intro', xPx: 0, widthPx: 50 },
			{ text: '', xPx: 50, widthPx: 130, leader: 'dot' },
			{ text: '12', xPx: 180, widthPx: 20 },
		]);
	});

	it('centers and decimal-aligns text at their stops', () => {
		const centered = layoutParagraph(
			paragraph('\tabcd', { tabStops: [{ posPx: 100, align: 'center' }] }),
			200,
			measurer,
			noop,
		);
		expect(centered.lines[0].fragments[1].xPx).toBe(80);
		const decimal = layoutParagraph(
			paragraph('\t123.45', { tabStops: [{ posPx: 100, align: 'decimal' }] }),
			200,
			measurer,
			noop,
		);
		expect(decimal.lines[0].fragments[1].xPx).toBe(70);
	});

	it('lays out list labels in the hanging indent from the numbering definition', () => {
		const model: DocumentModel = {
			blocks: [
				{ type: 'paragraph', id: 'l', runs: [{ text: 'Item' }], numbering: { numId: 1, level: 0 } },
			],
			page: {
				width: 816,
				height: 1056,
				marginTop: 96,
				marginRight: 96,
				marginBottom: 96,
				marginLeft: 96,
			},
			warnings: [],
			numberingCatalog: {
				abstractNums: {
					'0': {
						id: '0',
						levels: {
							0: {
								level: 0,
								start: 1,
								numFmt: 'bullet',
								lvlText: '',
								indentLeftTwips: 720,
								hangingTwips: 360,
							},
						},
					},
				},
				nums: { '1': { id: '1', abstractNumId: '0', overrides: {} } },
			} as never,
		};
		const [adapted] = adaptDocumentModel(model).sections[0].blocks as LayoutParagraph[];
		expect(adapted.runs[0]).toMatchObject({ text: '•\t', synthetic: true });
		expect(adapted).toMatchObject({ indentLeftTwips: 720, hangingTwips: 360 });
		const result = layoutParagraph(adapted, 600, measurer, noop);
		// The bullet starts at 24px (720 - 360 twips); its tab reaches the 48px indent.
		expect(fragmentsOf(result)).toEqual([
			{ text: '•', xPx: 24, widthPx: 10 },
			{ text: '', xPx: 34, widthPx: 14 },
			{ text: 'Item', xPx: 48, widthPx: 40 },
		]);
		expect(result.lines[0].sourceStart).toBe(0);
	});
});
