import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe('track changes setting', () => {
	it('parses settings.xml w:trackChanges', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document ${NS}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		zip.file('word/settings.xml', `<w:settings ${NS}><w:trackChanges/></w:settings>`);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		expect(loaded.model.trackChanges).toBe(true);
	});

	it('creates settings.xml, its relationship and content type when enabling track changes for the first time', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document ${NS}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		zip.file(
			'[Content_Types].xml',
			'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
		);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(bytes);
		expect(loaded.model.trackChanges).toBeUndefined();
		const saved = await JSZip.loadAsync(await loaded.save({ ...loaded.model, trackChanges: true }));
		const settingsXml = await saved.file('word/settings.xml')?.async('string');
		expect(settingsXml).toContain('<w:trackRevisions/>');
		expect(settingsXml).not.toContain('trackChanges');
		const rels = await saved.file('word/_rels/document.xml.rels')?.async('string');
		expect(rels).toContain('relationships/settings');
		const contentTypes = await saved.file('[Content_Types].xml')?.async('string');
		expect(contentTypes).toContain('settings+xml');

		const reopened = await loadDocx(await saved.generateAsync({ type: 'uint8array' }));
		expect(reopened.model.trackChanges).toBe(true);
		const disabled = await JSZip.loadAsync(
			await reopened.save({ ...reopened.model, trackChanges: false }),
		);
		const disabledSettings = await disabled.file('word/settings.xml')?.async('string');
		expect(disabledSettings).not.toContain('trackRevisions');
	});

	it('returns original bytes for a no-op save even once settings.xml exists', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document ${NS}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		zip.file('word/settings.xml', `<w:settings ${NS}><w:trackChanges/></w:settings>`);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
	});

	it('reads the Word w:trackRevisions flag and inserts flags in schema order', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document ${NS}><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		zip.file(
			'word/settings.xml',
			`<w:settings ${NS}><w:zoom w:percent="100"/><w:trackRevisions/><w:defaultTabStop w:val="720"/><w:compat/></w:settings>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		expect(loaded.model.trackChanges).toBe(true);
		const saved = await JSZip.loadAsync(
			await loaded.save({ ...loaded.model, evenAndOddHeaders: true }),
		);
		const xml = await saved.file('word/settings.xml')!.async('string');
		expect(xml).toMatch(/<w:defaultTabStop[^>]*\/><w:evenAndOddHeaders\/><w:compat\/>/);
	});
});
