import { describe, expect, it } from 'vitest';
import type {
	DocumentModel,
	ParagraphStyleCatalog,
	SectionProperties,
} from '@christophervr/docx-core';
import { adaptDocumentModel } from './adapter.js';

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

	it('maps docx-core sections, breaks and pageBreakBefore into layout input', () => {
		const section = (overrides: Partial<SectionProperties>): SectionProperties => ({
			endsAtBlockId: 'p1',
			type: 'nextPage',
			pageWidthTwips: 6000,
			pageHeightTwips: 15840,
			orientation: 'portrait',
			marginTopTwips: 1440,
			marginRightTwips: 1440,
			marginBottomTwips: 1440,
			marginLeftTwips: 1440,
			columns: { count: 1, equalWidth: true },
			...overrides,
		});
		const model: DocumentModel = {
			...baseModel({
				blocks: [
					{ type: 'paragraph', id: 'p1', runs: [{ text: 'A' }, { text: '', break: 'page' }] },
					{ type: 'paragraph', id: 'p2', runs: [{ text: 'B' }], pageBreakBefore: true },
				],
			}),
			sections: [
				section({}),
				section({
					endsAtBlockId: 'p2',
					type: 'continuous',
					columns: { count: 2, spacingTwips: 180, equalWidth: true },
				}),
			],
		};
		const input = adaptDocumentModel(model);
		expect(input.sections).toHaveLength(2);
		expect(input.sections[0].page.widthPx).toBe(400);
		expect(input.sections[0].page.marginTopPx).toBe(96);
		expect(input.sections[0].blocks).toHaveLength(1);
		expect(input.sections[0].break).toBeUndefined();
		expect(input.sections[1].columns).toEqual({ count: 2, gapPx: 12 });
		expect(input.sections[1].break).toBe('continuous');
		const first = input.sections[0].blocks[0];
		if (first.kind === 'paragraph') expect(first.runs[1].breakAfter).toBe('page');
		const second = input.sections[1].blocks[0];
		if (second.kind === 'paragraph') expect(second.pageBreakBefore).toBe(true);
	});
});
