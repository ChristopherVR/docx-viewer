import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, type InlineImage, type Paragraph } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function blank(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const xmlOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string')) ?? '';

const newBox = (lines: string[], border = true): InlineImage => ({
	relId: '',
	partName: '',
	contentType: '',
	widthPx: 192,
	heightPx: 96,
	unsupported: 'Text box',
	textBoxText: lines,
	textBoxEditable: true,
	textBoxBorder: border,
});

describe('inserted text boxes', () => {
	it('writes an inline wps text box that reads back as an editable box', async () => {
		const loaded = await loadDocx(await blank());
		(loaded.model.blocks[0] as Paragraph).runs.push({
			text: '',
			image: newBox(['Hello', '', ' padded ']),
		});
		const saved = await loaded.save();
		const xml = await xmlOf(saved);
		expect(xml).toContain('<wp:inline');
		expect(xml).toContain('txBox="1"');
		expect(xml).toMatch(/<wp:extent cx="1828800" cy="914400"\/>/);
		expect(xml).toContain('<w:p/>');
		expect(xml).toContain('xml:space="preserve"> padded </w:t>');
		const reopened = await loadDocx(saved);
		const image = (reopened.model.blocks[0] as Paragraph).runs.find((run) => run.image)!.image!;
		expect(image).toMatchObject({
			unsupported: 'Text box',
			textBoxEditable: true,
			textBoxBorder: true,
			widthPx: 192,
			heightPx: 96,
		});
		expect(image.textBoxText).toEqual(['Hello', '', ' padded ']);
	});

	it('edits text, size and outline in place and keeps other XML when only other text changes', async () => {
		const first = await loadDocx(await blank());
		(first.model.blocks[0] as Paragraph).runs.push({ text: '', image: newBox(['One']) });
		const saved = await first.save();
		const loaded = await loadDocx(saved);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		paragraph.runs[0]!.text = 'Changed before';
		expect(await xmlOf(await loaded.save())).toContain('<w:t>One</w:t>');
		const box = paragraph.runs.find((run) => run.image)!.image!;
		box.textBoxText = ['Two', 'Lines'];
		box.widthPx = 288;
		box.textBoxBorder = false;
		const xml = await xmlOf(await loaded.save());
		expect(xml).toContain('<w:t>Two</w:t>');
		expect(xml).not.toContain('<w:t>One</w:t>');
		expect(xml).toContain('cx="2743200"');
		expect(xml).toMatch(/<a:ln><a:noFill\/><\/a:ln>/);
		expect(xml.match(/<wp:inline/g)).toHaveLength(1);
	});

	it('does not offer editing for a floating or formatted Word text box', async () => {
		const zip = new JSZip();
		const formatted = `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="Text Box 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:p>${formatted}<w:r><w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.blocks[0] as Paragraph;
		expect(paragraph.runs[0]!.image!.textBoxEditable).toBeUndefined();
		// Editing the text beside it no longer rejects the paragraph, and the box XML stays as written.
		paragraph.runs[1]!.text = 'later';
		const xml = await xmlOf(await loaded.save());
		expect(xml).toContain('<w:jc w:val="center"/>');
		expect(xml).toContain('later');
	});
});
