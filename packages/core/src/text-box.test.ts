import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, type Paragraph } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const box = `<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" behindDoc="0"><wp:extent cx="1905000" cy="952500"/><wp:docPr id="1" name="Text Box 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:txbx><w:txbxContent><w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>box</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>`;

describe('text boxes', () => {
	it('reads the text of a wps text box and keeps it byte-for-byte on save', async () => {
		const zip = new JSZip();
		const xml = `<w:document xmlns:w="${W}"><w:body><w:p>${box}<w:r><w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
		zip.file('word/document.xml', xml);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const image = (loaded.model.blocks[0] as Paragraph).runs[0]!.image!;
		expect(image).toMatchObject({ unsupported: 'Text box', widthPx: 200, heightPx: 100 });
		expect(image.textBoxText).toEqual(['Hello box', 'Second']);
		const saved = await (
			await JSZip.loadAsync(await loaded.save(loaded.model))
		)
			.file('word/document.xml')!
			.async('string');
		expect(saved).toContain('Hello </w:t>');
		expect(saved).toContain('mc:Fallback');
	});
});
