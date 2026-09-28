import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, resolveParagraphFormatting, type Paragraph } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('paragraph keep options', () => {
	it('parses and edits keepNext, keepLines, widowControl and contextualSpacing', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:keepNext/><w:widowControl w:val="0"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.blocks[0] as Paragraph;
		expect(paragraph).toMatchObject({ keepNext: true, widowControl: false });
		const edited: Paragraph = {
			...paragraph,
			keepNext: undefined,
			keepLines: true,
			contextualSpacing: true,
		};
		const xml = await (
			await JSZip.loadAsync(await loaded.save({ ...loaded.model, blocks: [edited] }))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain(
			'<w:pPr><w:keepLines/><w:widowControl w:val="0"/><w:contextualSpacing/></w:pPr>',
		);
	});

	it("resolves Word's heading styles as keep-with-next", () => {
		const model = createDocument();
		const heading: Paragraph = { type: 'paragraph', id: 'h', style: 'Heading1', runs: [] };
		expect(resolveParagraphFormatting(heading, model.paragraphStyles!)).toMatchObject({
			keepNext: true,
			keepLines: true,
		});
	});
});
