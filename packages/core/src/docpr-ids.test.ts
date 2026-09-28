import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx, type DocumentModel, type Paragraph } from './index.js';
import { DocPrIdAllocator } from './docpr-ids.js';

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const picture = (name: string) => ({
	text: '',
	image: {
		relId: '',
		partName: `word/media/${name}.png`,
		contentType: 'image/png',
		widthPx: 10,
		heightPx: 10,
	},
});
const media = (...names: string[]) =>
	new Map(
		names.map((name) => [`word/media/${name}.png`, { bytes: PNG, contentType: 'image/png' }]),
	);
const docPrIds = (xml: string) => [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1]);
const part = async (bytes: Uint8Array, name: string) =>
	(await JSZip.loadAsync(bytes)).file(name)!.async('string');

describe('package-level docPr ids', () => {
	it('allocates past ids reserved from any source', () => {
		const allocator = new DocPrIdAllocator();
		allocator.reserveFromXml('<wp:docPr id="4" name="a"/><wp:docPr descr="x" id="9"/>');
		expect(allocator.next()).toBe('10');
		expect(allocator.next()).toBe('11');
	});

	it('gives pictures in the body and in a footnote different ids', async () => {
		const model = createDocument();
		(model.blocks[0] as Paragraph).runs = [
			picture('a'),
			{ text: '', noteReference: { kind: 'footnote', id: '1' } },
		];
		model.footnotes = [
			{
				id: '1',
				blocks: [
					{ type: 'paragraph', id: 'n1', runs: [{ text: '', noteMark: 'footnote' }, picture('b')] },
				],
			},
		];
		const saved = await saveDocx(model, media('a', 'b'));
		const body = docPrIds(await part(saved, 'word/document.xml'));
		const note = docPrIds(await part(saved, 'word/footnotes.xml'));
		expect(body).toHaveLength(1);
		expect(note).toHaveLength(1);
		expect(new Set([...body, ...note]).size).toBe(2);
	});

	it('avoids ids already used by an untouched header part', async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
		const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
		);
		zip.file(
			'word/_rels/document.xml.rels',
			`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${r}/header" Target="header1.xml"/></Relationships>`,
		);
		zip.file(
			'word/header1.xml',
			`<w:hdr xmlns:w="${w}" xmlns:wp="${wp}"><w:p><w:r><w:drawing><wp:inline><wp:docPr id="7" name="Logo"/></wp:inline></w:drawing></w:r></w:p></w:hdr>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const model = structuredClone(loaded.model) as DocumentModel;
		(model.blocks[0] as Paragraph).runs.push(picture('c'));
		const saved = await loaded.save(model, media('c'));
		expect(docPrIds(await part(saved, 'word/document.xml'))).toEqual(['8']);
	});
});
