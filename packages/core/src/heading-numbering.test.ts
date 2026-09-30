import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { computeListLabels, resolveParagraphNumbering } from './numbering-format.js';
import type { Paragraph } from './model.js';
import {
	headingLabels,
	headingNumberingFixture,
} from './test-support/heading-numbering-fixture.js';

describe('heading-linked numbering', () => {
	it('resolves levels from pStyle without flattening style numbering into paragraphs', async () => {
		const bytes = await headingNumberingFixture();
		const loaded = await loadDocx(bytes);
		expect([...computeListLabels(loaded.model).values()].map((label) => label.text)).toEqual(
			headingLabels,
		);
		expect(
			loaded.model.blocks.every((block) => block.type === 'paragraph' && !block.numbering),
		).toBe(true);
		expect(await loaded.save()).toEqual(bytes);
	});

	it('preserves styles and numbering XML when editing heading text', async () => {
		const bytes = await headingNumberingFixture();
		const loaded = await loadDocx(bytes);
		const model = structuredClone(loaded.model);
		(model.blocks[1] as Paragraph).runs[0]!.text = 'Changed heading';
		const saved = await loaded.save(model);
		const before = await JSZip.loadAsync(bytes);
		const after = await JSZip.loadAsync(saved);
		for (const part of ['word/styles.xml', 'word/numbering.xml'])
			expect(await after.file(part)!.async('string')).toBe(
				await before.file(part)!.async('string'),
			);
		expect(
			[...computeListLabels((await loadDocx(saved)).model).values()].map((label) => label.text),
		).toEqual(headingLabels);
	});

	it('honors direct numbering, explicit removal and effective level overrides', async () => {
		const { model } = await loadDocx(await headingNumberingFixture());
		const paragraph = model.blocks[1] as Paragraph;
		const resolve = (p: Paragraph) =>
			resolveParagraphNumbering(p, model.paragraphStyles, model.numberingCatalog);
		expect(resolve(paragraph)).toEqual({ numId: 1, level: 1 });
		model.paragraphStyles!.styles['Heading2']!.numbering = { numId: 1, level: 7 };
		expect(resolve(paragraph)).toEqual({ numId: 1, level: 1 });
		expect(resolve({ ...paragraph, numbering: { numId: 1, level: 0 } })).toEqual({
			numId: 1,
			level: 0,
		});
		expect(resolve({ ...paragraph, numbering: { numId: 0, level: 0 } })).toBeUndefined();
		model.numberingCatalog!.nums['1']!.levelOverrides = {
			1: { lvl: { level: 1, start: 1, numFmt: 'decimal', lvlText: '%2.' } },
		};
		expect(resolve(paragraph)).toEqual({ numId: 1, level: 7 });
		model.paragraphStyles!.styles['Heading2']!.numbering = { numId: 0, level: 0 };
		expect(resolve(paragraph)).toBeUndefined();
	});
});
