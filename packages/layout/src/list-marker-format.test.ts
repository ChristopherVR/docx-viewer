import { expect, it } from 'vitest';
import { createDocument, createListDefinition, twips } from '@christophervr/docx-core';
import { adaptDocumentModel } from './adapter.js';

it('formats the Print Layout marker run from the level, not the paragraph text', () => {
	const model = createDocument();
	const { catalog, numId } = createListDefinition(undefined, [
		{
			level: 0,
			start: 1,
			numFmt: 'decimal',
			lvlText: '%1.',
			markerFormat: { fontFamily: 'Georgia', fontSizeHalfPoints: 28, bold: true, color: '#1f4e79' },
		},
	]);
	model.numberingCatalog = catalog;
	model.blocks = [
		{ type: 'paragraph', id: 'p', numbering: { numId, level: 0 }, runs: [{ text: 'Body' }] },
	];
	const paragraph = adaptDocumentModel(model).sections[0]!.blocks[0]!;
	if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph');
	const marker = paragraph.runs.find((run) => run.synthetic)!;
	expect(marker).toMatchObject({
		fontFamily: 'Georgia',
		fontSizePt: 14,
		bold: true,
		color: '#1f4e79',
	});
});

it('adds the level tab stop to the paragraph tab stops in order', () => {
	const model = createDocument();
	const { catalog, numId } = createListDefinition(undefined, [
		{ level: 0, start: 1, numFmt: 'decimal', lvlText: '%1.', tabStopTwips: twips(1440) },
	]);
	model.numberingCatalog = catalog;
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p',
			numbering: { numId, level: 0 },
			tabStops: [{ posTwips: twips(2880), align: 'left' }],
			runs: [{ text: 'Body' }],
		} as never,
	];
	const paragraph = adaptDocumentModel(model).sections[0]!.blocks[0]!;
	if (paragraph.kind !== 'paragraph') throw new Error('Missing paragraph');
	expect(paragraph.tabStops?.map((stop) => stop.posPx)).toEqual([96, 192]);
});
