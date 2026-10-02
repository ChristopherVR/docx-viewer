import { createDocument, DEFAULT_TABLE_BORDERS, type DocumentModel } from 'docx-core';

/** The demo's sample document, shown from the landing page. */
export function createSampleDocument(): DocumentModel {
	const sample = createDocument();
	sample.blocks = [
		{
			type: 'paragraph',
			id: 'title',
			runs: [{ text: 'Document title', fontFamily: 'Calibri', fontSize: 26, bold: true }],
		},
		{
			type: 'paragraph',
			id: 'subtitle',
			runs: [{ text: 'Add a subtitle or date here', fontSize: 10, color: '#666666' }],
		},
		{
			type: 'paragraph',
			id: 'intro',
			runs: [
				{
					text: 'Start with your first paragraph. Use the ribbon to set type, size, alignment, and page options as you work.',
				},
			],
		},
		{
			type: 'paragraph',
			id: 'body',
			runs: [
				{ text: 'A document workspace. ', bold: true },
				{
					text: 'Use the ribbon to format text, add a simple table, or adjust page settings. Open a DOCX file to continue working, or start a new document.',
				},
			],
		},
		{
			type: 'paragraph',
			id: 'quote',
			align: 'center',
			runs: [{ text: 'Clarity is a kindness to the reader.', italic: true, fontSize: 16 }],
		},
		{
			type: 'paragraph',
			id: 'heading',
			runs: [{ text: 'A small plan for a good first draft', bold: true }],
		},
		{
			type: 'table',
			id: 'table',
			borders: DEFAULT_TABLE_BORDERS,
			rows: [
				[
					{
						paragraphs: [
							{ type: 'paragraph', id: 'h1', runs: [{ text: 'START WITH', bold: true }] },
						],
					},
					{
						paragraphs: [
							{ type: 'paragraph', id: 'h2', runs: [{ text: 'MAKE IT COUNT', bold: true }] },
						],
					},
				],
				[
					{ paragraphs: [{ type: 'paragraph', id: 'c1', runs: [{ text: 'One clear idea' }] }] },
					{
						paragraphs: [
							{ type: 'paragraph', id: 'c2', runs: [{ text: 'Give each paragraph a purpose.' }] },
						],
					},
				],
				[
					{ paragraphs: [{ type: 'paragraph', id: 'c3', runs: [{ text: 'A thoughtful edit' }] }] },
					{
						paragraphs: [
							{ type: 'paragraph', id: 'c4', runs: [{ text: 'Read it once more, then share.' }] },
						],
					},
				],
			],
		},
	];
	return sample;
}
