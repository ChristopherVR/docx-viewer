import type { DocumentModel } from '@christophervr/docx-core';
import { describe, expect, it } from 'vitest';
import { adaptDocumentModel } from './adapter.js';

describe('drop cap folding', () => {
	it('lays the initial out at the start of the next paragraph', () => {
		const model: DocumentModel = {
			blocks: [
				{
					type: 'paragraph',
					id: 'cap',
					runs: [{ text: 'O', fontSize: 55 }],
					dropCap: { style: 'drop', lines: 3 },
				},
				{ type: 'paragraph', id: 'body', runs: [{ text: 'nce' }] },
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
		};
		const blocks = adaptDocumentModel(model).sections[0]!.blocks;
		expect(blocks).toHaveLength(1);
		const body = blocks[0] as { runs: Array<{ text?: string }> };
		expect(body.runs.map((run) => run.text)).toEqual(['O', 'nce']);
	});
});
