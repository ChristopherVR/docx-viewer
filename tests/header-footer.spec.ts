import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { fileInput, fileNameLabel, saveButton } from './helpers';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function headerDocx(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${r}/header" Target="header1.xml"/></Relationships>`,
	);
	zip.file(
		'word/header1.xml',
		`<w:hdr xmlns:w="${w}"><w:p><w:r><w:t>Company</w:t></w:r></w:p></w:hdr>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

test('edits a header in place by double-clicking it and saves the header part', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'header.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await headerDocx(),
	});
	await expect(fileNameLabel(page)).toHaveText('header.docx');
	const editor = page.locator('docx-editor');
	const header = editor.locator('.dve-header .dve-header-footer-slot');
	await expect(header).toContainText('Company');
	await header.dblclick();
	const headerEditor = header.locator('.dve-header-footer-editor .ProseMirror');
	await expect(headerEditor).toBeFocused();
	await page.keyboard.press('End');
	await page.keyboard.type(' Confidential');
	await page.keyboard.press('Escape');
	await expect(header).toContainText('Company Confidential');
	await expect(editor.locator('.dve-save-state')).toHaveText('Unsaved changes');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	expect(await zip.file('word/header1.xml')!.async('string')).toContain('Company Confidential');
	expect(await zip.file('word/document.xml')!.async('string')).toContain('Body text');
	await expect(editor.locator('.dve-save-state')).toHaveText('Saved to this PC');
	expect(errors).toEqual([]);
});
