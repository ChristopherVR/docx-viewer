import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import type { Paragraph } from './model.js';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p><w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/numbering.xml',
		`<?xml version="1.0"?><w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('numbering-write', () => {
	it('changes the level in place without touching numId', async () => {
		const loaded = await loadDocx(await fixture());
		const paragraph = loaded.model.blocks[0] as Paragraph;
		paragraph.numbering = { numId: 1, level: 0 };
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).toContain('<w:ilvl w:val="0"');
		expect(xml).toContain('<w:numId w:val="1"');
	});

	it('adds numPr to a previously plain paragraph', async () => {
		const loaded = await loadDocx(await fixture());
		const paragraph = loaded.model.blocks[1] as Paragraph;
		paragraph.numbering = { numId: 1, level: 0 };
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('word/document.xml')?.async('string')) ?? '';
		expect(xml.match(/<w:numPr>/g)).toHaveLength(2);
	});

	it('removes numPr entirely when numbering is cleared', async () => {
		const loaded = await loadDocx(await fixture());
		const paragraph = loaded.model.blocks[0] as Paragraph;
		delete paragraph.numbering;
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).not.toContain('<w:numPr>');
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		expect((reopened.model.blocks[0] as Paragraph).numbering).toBeUndefined();
	});

	it('writes an explicit numId=0 override rather than an absent numPr', async () => {
		const loaded = await loadDocx(await fixture());
		const paragraph = loaded.model.blocks[0] as Paragraph;
		paragraph.numbering = { numId: 0, level: 0 };
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).toContain('<w:numId w:val="0"');
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		expect((reopened.model.blocks[0] as Paragraph).numbering).toEqual({ numId: 0, level: 0 });
	});
});
