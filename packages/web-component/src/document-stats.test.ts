import { createDocument, type DocumentModel } from 'docx-core';
import { describe, expect, it } from 'vitest';
import { documentStats, plainText } from './document-stats';

function model(): DocumentModel {
	const doc = createDocument();
	doc.blocks = [
		{ type: 'paragraph', id: 'a', runs: [{ text: 'Hello ' }, { text: 'world' }] },
		{
			type: 'paragraph',
			id: 'b',
			runs: [
				{ text: '', fieldChar: 'begin' },
				{ text: '', fieldCode: ' PAGE ' },
				{ text: '', fieldChar: 'separate' },
				{ text: '7' },
				{ text: '', fieldChar: 'end' },
				{ text: 'gone', revision: { kind: 'delete', id: '1', author: 'x' } as never },
				{ text: 'kept', revision: { kind: 'insert', id: '2', author: 'x' } as never },
			],
		},
		{
			type: 'table',
			id: 't',
			rows: [
				[
					{ paragraphs: [{ type: 'paragraph', id: 'c1', runs: [{ text: 'A' }] }] },
					{ paragraphs: [{ type: 'paragraph', id: 'c2', runs: [{ text: 'B' }, { text: '!' }] }] },
				],
			],
		} as never,
	];
	doc.comments = [{ id: '1', author: 'Ann', text: 'hi' }];
	return doc;
}

describe('documentStats', () => {
	it('counts blocks, characters, comments and revisions', () => {
		const stats = documentStats(model());
		expect(stats).toMatchObject({
			paragraphs: 4,
			tables: 1,
			sections: 1,
			comments: 1,
			revisions: 2,
		});
		// "Hello world" + "7kept" + "A" + "B!"
		expect(stats.charactersWithSpaces).toBe('Hello world7keptAB!'.length);
		expect(stats.characters).toBe('Helloworld7keptAB!'.length);
	});

	it('handles an empty document', () => {
		expect(documentStats(createDocument())).toMatchObject({
			characters: 0,
			comments: 0,
			revisions: 0,
		});
	});
});

describe('plainText', () => {
	it('writes one line per paragraph and tab-separated table rows, without deleted text or fields', () => {
		expect(plainText(model())).toBe('Hello world\n7kept\nA\tB!');
	});
});
