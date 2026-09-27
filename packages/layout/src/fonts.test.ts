import { describe, expect, it } from 'vitest';
import type { DocumentModel } from '@christophervr/docx-core';
import { adaptDocumentModel } from './adapter.js';
import { cssFontStack } from './fonts.js';
import type { LayoutParagraph } from './input.js';

describe('font fidelity', () => {
	it('adds metric-compatible substitutes and a generic family', () => {
		expect(cssFontStack('Calibri')).toBe('"Calibri", "Carlito", sans-serif');
		expect(cssFontStack('Times New Roman')).toBe(
			'"Times New Roman", "Tinos", "Liberation Serif", serif',
		);
		expect(cssFontStack('Courier New')).toBe(
			'"Courier New", "Cousine", "Liberation Mono", monospace',
		);
		expect(cssFontStack('Segoe UI')).toBe('"Segoe UI", sans-serif');
		expect(cssFontStack(undefined)).toBe('"Calibri", "Carlito", sans-serif');
	});

	it('measures runs with formatting inherited from defaults, styles and the theme', () => {
		const model: DocumentModel = {
			blocks: [
				{
					type: 'paragraph',
					id: 'h',
					style: 'Heading1',
					runs: [{ text: 'Title' }, { text: 'plain', bold: true }],
				},
				{
					type: 'paragraph',
					id: 'b',
					runs: [{ text: 'secret', vanish: true } as never, { text: 'caps', caps: true } as never],
				},
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
			paragraphStyles: {
				docDefaults: {},
				warnings: [],
				styles: { Heading1: { id: 'Heading1', name: 'heading 1', formatting: {} } },
			},
			characterStyles: {
				docDefaults: { fontSize: 11, fontTheme: { ascii: 'minor' } },
				warnings: [],
				styles: {
					Heading1: {
						id: 'Heading1',
						type: 'paragraph',
						formatting: {
							fontSize: 16,
							bold: true,
							fontTheme: { ascii: 'major' },
							color: '#2F5496',
						},
					},
				},
			},
			theme: {
				fonts: { major: { latin: 'Calibri Light' }, minor: { latin: 'Calibri' } },
			} as never,
		};
		const [heading, body] = adaptDocumentModel(model).sections[0].blocks as LayoutParagraph[];
		expect(heading.runs[0]).toMatchObject({
			text: 'Title',
			fontSizePt: 16,
			bold: true,
			fontFamily: 'Calibri Light',
			color: '#2F5496',
		});
		// Direct bold on top of a bold style toggles it off, as in Word.
		expect(heading.runs[1].bold).toBeFalsy();
		expect(body.runs.map((run) => run.text)).toEqual(['', 'CAPS']);
		expect(body.runs[1]).toMatchObject({ fontFamily: 'Calibri', fontSizePt: 11 });
	});
});
