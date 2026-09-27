import { describe, expect, it } from 'vitest';
import type { DocumentModel, ParagraphStyleCatalog } from '@christophervr/docx-core';
import { adaptDocumentModel, type ModelWithSections } from './adapter.js';

function baseModel(overrides: Partial<DocumentModel> = {}): DocumentModel {
	return {
		blocks: [{ type: 'paragraph', id: 'p1', runs: [{ text: 'Hello' }] }],
		page: {
			width: 816,
			height: 1056,
			marginTop: 96,
			marginRight: 96,
			marginBottom: 96,
			marginLeft: 96,
		},
		warnings: [],
		...overrides,
	};
}

describe('adaptDocumentModel', () => {
	it('produces a single section from model.page when no sections hint is present', () => {
		const input = adaptDocumentModel(baseModel());
		expect(input.sections).toHaveLength(1);
		expect(input.sections[0].page).toMatchObject({
			widthPx: 816,
			heightPx: 1056,
			marginTopPx: 96,
			marginRightPx: 96,
			marginBottomPx: 96,
			marginLeftPx: 96,
		});
		expect(input.sections[0].blocks).toHaveLength(1);
	});

	it('converts table blocks with their nested cell paragraphs', () => {
		const model = baseModel({
			blocks: [
				{
					type: 'table',
					id: 't1',
					rows: [[{ paragraphs: [{ type: 'paragraph', id: 'c1', runs: [{ text: 'Cell' }] }] }]],
				},
			],
		});
		const input = adaptDocumentModel(model);
		const table = input.sections[0].blocks[0];
		expect(table.kind).toBe('table');
		if (table.kind === 'table')
			expect(table.rows[0].cells[0].paragraphs[0].runs[0].text).toBe('Cell');
	});

	it('reports honest approximations for features the model does not yet expose', () => {
		const notes: string[] = [];
		adaptDocumentModel(baseModel(), (message) => notes.push(message));
		expect(notes.some((n) => n.includes('keepNext'))).toBe(true);
		// Deduplicated even though only one paragraph triggered it.
		expect(notes.filter((n) => n.includes('keepNext'))).toHaveLength(1);
	});

	it('resolves style-based formatting through resolveParagraphFormatting when a catalog is present', () => {
		const catalog: ParagraphStyleCatalog = {
			docDefaults: {},
			styles: { Body: { id: 'Body', formatting: { align: 'center' } } },
			warnings: [],
		};
		const model = baseModel({
			blocks: [{ type: 'paragraph', id: 'p1', runs: [{ text: 'x' }], style: 'Body' }],
			paragraphStyles: catalog,
		});
		const input = adaptDocumentModel(model);
		const paragraph = input.sections[0].blocks[0];
		expect(paragraph.kind).toBe('paragraph');
		if (paragraph.kind === 'paragraph') expect(paragraph.align).toBe('center');
	});

	it('plugs in a future sections hint when the model carries one', () => {
		const model: ModelWithSections = {
			...baseModel({
				blocks: [
					{ type: 'paragraph', id: 'p1', runs: [{ text: 'A' }] },
					{ type: 'paragraph', id: 'p2', runs: [{ text: 'B' }] },
				],
			}),
			sections: [
				{ startBlockIndex: 0, endBlockIndex: 1, page: { width: 400 } },
				{
					startBlockIndex: 1,
					endBlockIndex: 2,
					columns: { count: 2, gapPx: 12 },
					break: 'continuous',
				},
			],
		};
		const input = adaptDocumentModel(model);
		expect(input.sections).toHaveLength(2);
		expect(input.sections[0].page.widthPx).toBe(400);
		expect(input.sections[0].blocks).toHaveLength(1);
		expect(input.sections[1].columns).toEqual({ count: 2, gapPx: 12 });
		expect(input.sections[1].break).toBe('continuous');
	});
});
