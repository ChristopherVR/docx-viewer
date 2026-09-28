// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph } from '@christophervr/docx-core';
import { NodeSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS = `xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;

async function pictureDocx(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body><w:p><w:r><w:t>Logo </w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="Picture 1" descr="Old"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
	);
	zip.file('word/media/image1.png', new Uint8Array([137, 80, 78, 71]));
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/></Types>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('Format picture', () => {
	afterEach(() => document.body.replaceChildren());

	it('edits alt text and size and writes them back into the existing drawing', async () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		await editor.load(await pictureDocx());
		const view = (editor as unknown as { view: EditorView }).view;
		const root = editor.shadowRoot!;
		const format = root.querySelector<HTMLButtonElement>('[aria-label="Format picture"]')!;
		expect(format.disabled).toBe(true);
		let imagePos = -1;
		view.state.doc.descendants((node, pos) => {
			if (node.type.name === 'image') imagePos = pos;
		});
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, imagePos)));
		expect(format.disabled).toBe(false);
		format.click();
		const dialog = root.querySelector<HTMLElement>('.dve-picture-dialog')!;
		expect(dialog.hidden).toBe(false);
		const width = dialog.querySelector<HTMLInputElement>('[aria-label="Width"]')!;
		const height = dialog.querySelector<HTMLInputElement>('[aria-label="Height"]')!;
		expect([width.value, height.value]).toEqual(['100', '50']);
		width.value = '200';
		width.dispatchEvent(new Event('input'));
		expect(height.value).toBe('100');
		dialog.querySelector<HTMLTextAreaElement>('[aria-label="Alt text"]')!.value = 'Company logo';
		[...dialog.querySelectorAll('button')].find((button) => button.textContent === 'OK')!.click();
		const image = (editor.documentModel!.blocks[0] as Paragraph).runs[1].image;
		expect(image).toMatchObject({ widthPx: 200, heightPx: 100, altText: 'Company logo' });
		const saved = await editor.saveBytes();
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).toContain('cx="1905000" cy="952500"');
		expect(xml).toContain('descr="Company logo"');
		expect((await loadDocx(saved)).media?.get('word/media/image1.png')).toBeDefined();
	});
});
