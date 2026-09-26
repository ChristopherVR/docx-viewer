import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { isWordHighlightToken, loadDocx, saveDocx, WORD_HIGHLIGHT_TOKENS } from './index.js';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function packageWith(xml: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('word/document.xml', xml);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('Word highlight tokens', () => {
	it('exports the canonical Word highlight tokens and validates membership', () => {
		expect(WORD_HIGHLIGHT_TOKENS).toEqual([
			'black',
			'blue',
			'cyan',
			'green',
			'magenta',
			'red',
			'yellow',
			'white',
			'darkBlue',
			'darkCyan',
			'darkGreen',
			'darkMagenta',
			'darkRed',
			'darkYellow',
			'darkGray',
			'lightGray',
			'none',
		]);
		for (const token of WORD_HIGHLIGHT_TOKENS) expect(isWordHighlightToken(token)).toBe(true);
		expect(isWordHighlightToken('gray25')).toBe(false);
		expect(isWordHighlightToken('brightGreen')).toBe(false);
	});

	it('preserves unsupported source highlight tokens unchanged and rejects new invalid values', async () => {
		const original = await packageWith(
			`<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:rPr><w:highlight w:val="gray25"/></w:rPr><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(original);
		expect(await loaded.save()).toEqual(original);
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraph.runs[0].highlight).toBe('gray25');
		paragraph.runs[0].text = 'Edited';
		const preserved = await JSZip.loadAsync(await loaded.save());
		expect(await preserved.file('word/document.xml')?.async('string')).toContain('w:val="gray25"');
		paragraph.runs[0].highlight = 'brightGreen';
		await expect(loaded.save()).rejects.toThrow('Unsupported Word highlight token: brightGreen');
	});

	it('rejects invalid highlight tokens in new documents', async () => {
		await expect(
			saveDocx({
				blocks: [{ type: 'paragraph', id: 'p', runs: [{ text: 'New', highlight: 'gray25' }] }],
				page: {
					width: 816,
					height: 1056,
					marginTop: 96,
					marginRight: 96,
					marginBottom: 96,
					marginLeft: 96,
				},
				warnings: [],
			}),
		).rejects.toThrow('Unsupported Word highlight token: gray25');
	});

	it('rejects run splitting when it could drop unsupported run properties', async () => {
		const original = await packageWith(
			`<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:rPr><w:shd w:fill="ABCDEF"/></w:rPr><w:t>Formatted</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(original);
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		const originalRun = paragraph.runs[0];
		paragraph.runs.splice(0, 1, { text: 'Format' }, { text: 'ted', bold: true });
		await expect(loaded.save()).rejects.toThrow(
			'changing run boundaries could drop unsupported run properties',
		);
		const untouched = await loadDocx(original);
		expect(await untouched.save()).toEqual(original);
		expect(originalRun.text).toBe('Formatted');
	});

	it('treats unmodeled attributes inside known run properties as opaque', async () => {
		const original = await packageWith(
			`<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Noto"/></w:rPr><w:t>Formatted</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(original);
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		paragraph.runs.splice(0, 1, { text: 'Format' }, { text: 'ted', bold: true });
		await expect(loaded.save()).rejects.toThrow(
			'changing run boundaries could drop unsupported run properties',
		);
	});
});
