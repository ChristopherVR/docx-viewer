import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CORE = `<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Old</dc:title><dc:creator>Ann</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created></cp:coreProperties>`;

async function fixture(withCore: boolean): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	if (withCore) zip.file('docProps/core.xml', CORE);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('document properties', () => {
	it('reads title and author and returns the package unchanged when untouched', async () => {
		const bytes = await fixture(true);
		const loaded = await loadDocx(bytes);
		expect(loaded.model.properties).toEqual({ title: 'Old', creator: 'Ann' });
		expect(await loaded.save()).toEqual(bytes);
	});

	it('edits properties in place, keeping unrelated elements', async () => {
		const loaded = await loadDocx(await fixture(true));
		loaded.model.properties = { title: 'New & <better>', keywords: 'a, b' };
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('docProps/core.xml')?.async('string')) ?? '';
		expect(xml).toContain('2024-01-02T03:04:05Z');
		expect(xml).not.toContain('Ann');
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		expect(reopened.model.properties).toEqual({ title: 'New & <better>', keywords: 'a, b' });
	});

	it('creates the part, content type and root relationship when missing', async () => {
		const loaded = await loadDocx(await fixture(false));
		loaded.model.properties = { creator: 'Bo', subject: 'S' };
		const output = await JSZip.loadAsync(await loaded.save());
		expect(await output.file('[Content_Types].xml')?.async('string')).toContain(
			'/docProps/core.xml',
		);
		expect(await output.file('_rels/.rels')?.async('string')).toContain('core-properties');
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		expect(reopened.model.properties).toEqual({ creator: 'Bo', subject: 'S' });
	});
});
