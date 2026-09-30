import { expect, it } from 'vitest';
import {
	createDocument,
	createListDefinition,
	signedTwips,
	twips,
	type Paragraph,
} from '@christophervr/docx-core';
import { adaptDocumentModel } from './adapter.js';
import { layoutParagraph } from './paragraph-layout.js';
import type { TextMeasurer } from './measure.js';

const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };
const noop = () => {};

function adapted(alignment: 'left' | 'center' | 'right', suffix: 'tab' | 'space' | 'none' = 'tab') {
	const model = createDocument();
	const { catalog, numId } = createListDefinition(undefined, [
		{
			level: 0,
			start: 12,
			numFmt: 'decimal',
			lvlText: 'Part %1.',
			lvlJc: alignment,
			suffix,
			indentLeftTwips: signedTwips(1440),
			hangingTwips: twips(1080),
		},
	]);
	model.numberingCatalog = catalog;
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p',
			numbering: { numId, level: 0 },
			runs: [{ text: 'Body text wrapping after several words' }],
		},
	];
	const paragraph = adaptDocumentModel(model).sections[0]!.blocks[0]!;
	if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph');
	return { model, paragraph };
}

for (const [alignment, offset] of [
	['left', 0],
	['center', 40],
	['right', 80],
] as const) {
	it(`aligns the whole ${alignment} marker at its number position, with text at the next tab stop`, () => {
		const { paragraph } = adapted(alignment);
		const result = layoutParagraph(paragraph, 500, measurer, noop);
		const [marker, , body] = result.lines[0]!.fragments;
		expect(marker).toMatchObject({ text: 'Part 12.', xPx: 24 - offset, widthPx: 80 });
		// Left-aligned marker exceeds the text indent, so its tab advances to 144px;
		// centered and right markers end before 96px and use the hanging stop.
		expect(body!.xPx).toBe(alignment === 'left' ? 144 : 96);
	});
}

for (const [suffix, gap] of [
	['space', 10],
	['none', 0],
] as const) {
	it(`follows an aligned marker with ${suffix} and wraps continuation text at the text indent`, () => {
		const { paragraph } = adapted('right', suffix);
		const result = layoutParagraph(paragraph, 220, measurer, noop);
		const first = result.lines[0]!.fragments;
		expect(first[0]).toMatchObject({ text: 'Part 12.', xPx: -56 });
		expect(first.find((fragment) => fragment.text === 'Body')!.xPx).toBe(24 + gap);
		expect(result.lines[1]!.fragments[0]!.xPx).toBe(96);
	});
}

it('does not repeat generated markers on a continuation page or change source text positions', () => {
	const { model, paragraph } = adapted('right');
	const whole = layoutParagraph(paragraph, 220, measurer, noop);
	const continued = layoutParagraph(
		paragraph,
		220,
		measurer,
		noop,
		undefined,
		whole.lines[0]!.nextToken,
	);
	expect(
		continued.lines
			.flatMap((line) => line.fragments)
			.some((fragment) => fragment.text === 'Part 12.'),
	).toBe(false);
	expect(continued.lines[0]!.fragments[0]!.xPx).toBe(96);
	expect(whole.lines[0]!.sourceStart).toBe(0);
	expect((model.blocks[0] as Paragraph).runs[0]!.text).toBe(
		'Body text wrapping after several words',
	);
});
