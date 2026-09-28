import { describe, expect, it } from 'vitest';
import type { DocumentModel } from '@christophervr/docx-core';
import { adaptDocumentModel } from './adapter.js';
import { cssFontStack } from './fonts.js';
import type { LayoutParagraph } from './input.js';
import { at } from './__tests__/helpers.js';

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
		const heading = at(at(adaptDocumentModel(model).sections, 0).blocks as LayoutParagraph[], 0);
		const body = at(at(adaptDocumentModel(model).sections, 0).blocks as LayoutParagraph[], 1);
		expect(at(heading.runs, 0)).toMatchObject({
			text: 'Title',
			fontSizePt: 16,
			bold: true,
			fontFamily: 'Calibri Light',
			color: '#2F5496',
		});
		// Direct formatting is absolute: bold on a bold style stays bold (ECMA-376 §17.7.3).
		expect(at(heading.runs, 1).bold).toBe(true);
		expect(body.runs.map((run) => run.text)).toEqual(['', 'CAPS']);
		expect(at(body.runs, 1)).toMatchObject({ fontFamily: 'Calibri', fontSizePt: 11 });
	});
});
