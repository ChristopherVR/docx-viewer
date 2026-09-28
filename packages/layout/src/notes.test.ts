import { describe, expect, it } from 'vitest';
import type { DocumentModel, Paragraph } from '@christophervr/docx-core';
import { layoutDocumentModel } from './layout.js';
import type { LayoutParagraphBox } from './result.js';
import type { TextMeasurer } from './measure.js';
import { at } from './__tests__/helpers.js';

// 10px per character; every line is 20px tall.
const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };

function model(blocks: Paragraph[], extra: Partial<DocumentModel> = {}): DocumentModel {
	return {
		blocks,
		// A 200px-tall content area: ten 20px lines.
		page: {
			width: 600,
			height: 400,
			marginTop: 100,
			marginRight: 100,
			marginBottom: 100,
			marginLeft: 100,
		},
		warnings: [],
		footnotes: [
			{
				id: '1',
				blocks: [
					{
						type: 'paragraph',
						id: 'n1',
						runs: [{ text: '', noteMark: 'footnote' }, { text: ' Source' }],
					},
				],
			},
		],
		endnotes: [
			{
				id: '7',
				blocks: [
					{
						type: 'paragraph',
						id: 'e7',
						runs: [{ text: '', noteMark: 'endnote' }, { text: ' Later' }],
					},
				],
			},
		],
		...extra,
	};
}
const withNote = (id: string, text: string): Paragraph => ({
	type: 'paragraph',
	id,
	runs: [{ text }, { text: '', noteReference: { kind: 'footnote', id: '1' } }],
});
const lines = (id: string, count: number): Paragraph => ({
	type: 'paragraph',
	id,
	runs: [{ text: Array.from({ length: count }, () => 'x').join('\n') }],
});

describe('notes in Print Layout', () => {
	it('shows reference numbers raised and puts the footnote at the bottom of its page', () => {
		const result = layoutDocumentModel(model([withNote('p', 'Claim')]), measurer);
		const page = at(result.pages, 0);
		const body = at(at(page.columns, 0).blocks, 0) as LayoutParagraphBox;
		expect(at(body.lines, 0).fragments.map((f) => [f.text, f.script])).toEqual([
			['Claim', undefined],
			['1', 'super'],
		]);
		expect(page.footnotes).toHaveLength(1);
		const note = at(page.footnotes, 0);
		expect(note).toMatchObject({ id: '1', yPx: 0 });
		// This measurer gives small text a full-size ascent, so the raised mark adds a little height.
		expect(note.heightPx).toBeGreaterThanOrEqual(20);
		expect(at(at(at(note.paragraphs, 0).lines, 0).fragments, 0)).toMatchObject({
			text: '1',
			script: 'super',
		});
	});

	it('keeps a paragraph with its footnote: both move to the next page when they do not fit', () => {
		// Nine filler lines leave one line, but the paragraph also needs its note and separator.
		const result = layoutDocumentModel(model([lines('f', 9), withNote('p', 'Claim')]), measurer);
		expect(result.pages).toHaveLength(2);
		expect(at(result.pages, 0).footnotes).toBeUndefined();
		expect(at(at(at(result.pages, 1).columns, 0).blocks, 0).blockId).toBe('p');
		expect(at(at(result.pages, 1).footnotes, 0).id).toBe('1');
	});

	it('places endnotes after the last paragraph, in reference order', () => {
		const body: Paragraph = {
			type: 'paragraph',
			id: 'p',
			runs: [{ text: 'End' }, { text: '', noteReference: { kind: 'endnote', id: '7' } }],
		};
		const result = layoutDocumentModel(model([body]), measurer);
		const blocks = at(at(result.pages, 0).columns, 0).blocks as LayoutParagraphBox[];
		expect(blocks.map((block) => block.blockId)).toEqual(['p', 'e7']);
		expect(at(at(at(blocks, 0).lines, 0).fragments, 1)).toMatchObject({
			text: 'i',
			script: 'super',
		});
		expect(at(at(at(blocks, 1).lines, 0).fragments, 0)).toMatchObject({
			text: 'i',
			script: 'super',
		});
	});
});
