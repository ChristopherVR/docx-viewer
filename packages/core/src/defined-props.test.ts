import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { definedProps } from './defined-props.js';
import { computeListLabels } from './numbering-format.js';
import type { DocumentModel } from './model.js';
import type { NumberingCatalog } from './numbering-model.js';
import { loadDocx } from './parse.js';

describe('definedProps', () => {
	it('drops undefined values but keeps other falsy values', () => {
		const result = definedProps({ a: undefined, b: 0, c: false, d: '', e: null, f: 'x' });
		expect(Object.keys(result)).toEqual(['b', 'c', 'd', 'e', 'f']);
		expect(result).toStrictEqual({ b: 0, c: false, d: '', e: null, f: 'x' });
	});
});

describe('absent optional keys', () => {
	it('list labels omit unset indents rather than carrying undefined keys', () => {
		const numbering: NumberingCatalog = {
			abstractNums: {
				'0': {
					id: '0',
					levels: { 0: { level: 0, start: 1, numFmt: 'decimal', lvlText: '%1.', suffix: 'tab' } },
				},
			},
			nums: { '1': { id: '1', abstractNumId: '0' } },
			warnings: [],
		};
		const model: DocumentModel = {
			blocks: [
				{ type: 'paragraph', id: 'p', runs: [{ text: 'a' }], numbering: { numId: 1, level: 0 } },
			],
			page: { width: 1, height: 1, marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0 },
			warnings: [],
			numberingCatalog: numbering,
		};
		expect(computeListLabels(model).get('p')).toStrictEqual({
			numId: '1',
			level: 0,
			text: '1.',
			suffix: 'tab',
		});
	});

	it('parsed plain runs carry only their text key', async () => {
		const zip = new JSZip();
		const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.blocks[0];
		if (paragraph?.type !== 'paragraph') throw new Error('Expected a paragraph');
		expect(Object.keys(paragraph.runs[0] ?? {})).toEqual(['text']);
	});
});
