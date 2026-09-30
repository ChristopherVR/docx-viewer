import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const CONTENT_TYPES =
	'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';

async function packageWith(settings: string | null): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', CONTENT_TYPES);
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	if (settings !== null)
		zip.file('word/settings.xml', `<w:settings ${NS}>${settings}</w:settings>`);
	return zip.generateAsync({ type: 'uint8array' });
}
const settingsOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/settings.xml')?.async('string')) ?? '';

describe('automatic hyphenation', () => {
	it('reads w:autoHyphenation, honouring an explicit off', async () => {
		expect((await loadDocx(await packageWith('<w:autoHyphenation/>'))).model.autoHyphenation).toBe(
			true,
		);
		expect(
			(await loadDocx(await packageWith('<w:autoHyphenation w:val="0"/>'))).model.autoHyphenation,
		).toBeUndefined();
		expect((await loadDocx(await packageWith(null))).model.autoHyphenation).toBeUndefined();
	});

	it('writes the flag in schema order and removes it again', async () => {
		const loaded = await loadDocx(
			await packageWith('<w:zoom w:percent="100"/><w:evenAndOddHeaders/>'),
		);
		const on = await loaded.save({ ...loaded.model, autoHyphenation: true });
		const xml = await settingsOf(on);
		expect(xml).toMatch(/<w:zoom[^>]*\/><w:autoHyphenation\/><w:evenAndOddHeaders\/>/);
		const again = await loadDocx(on);
		const off = { ...again.model };
		delete off.autoHyphenation;
		expect(await settingsOf(await again.save(off))).not.toContain('autoHyphenation');
	});
});
