import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx, type Paragraph } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('toggle properties', () => {
	it('keeps an explicit off that cancels a style when the paragraph is edited', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:b w:val="0"/><w:u w:val="none"/></w:rPr><w:t>plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.blocks[0] as Paragraph;
		expect(paragraph.runs[0]).toMatchObject({ bold: false, underline: false });
		const edited = { ...paragraph, runs: [{ ...paragraph.runs[0], text: 'still plain' }] };
		const xml = await (
			await JSZip.loadAsync(await loaded.save({ ...loaded.model, blocks: [edited] }))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain(
			'<w:rPr><w:b w:val="0"/><w:u w:val="none"/></w:rPr><w:t>still plain</w:t>',
		);
	});

	it('writes run properties in schema order', async () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [
					{
						text: 'x',
						fontSize: 14,
						color: '#FF0000',
						bold: true,
						style: 'Hyperlink',
						underline: true,
						italic: false,
					},
				],
			},
		];
		const xml = await (
			await JSZip.loadAsync(await saveDocx(model))
		)
			.file('word/document.xml')!
			.async('string');
		const order = [...xml.matchAll(/<w:(rStyle|b|i|color|sz|u)\b/g)].map((match) => match[1]);
		expect(order).toEqual(['rStyle', 'b', 'i', 'color', 'sz', 'u']);
	});
});
