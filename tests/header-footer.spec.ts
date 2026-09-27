import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { fileInput, fileNameLabel, newDocument, saveButton } from './helpers';

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

async function footnoteDocx(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>Claim</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${r}/footnotes" Target="footnotes.xml"/></Relationships>`,
	);
	zip.file(
		'word/footnotes.xml',
		`<w:footnotes xmlns:w="${w}"><w:footnote w:id="0" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Source</w:t></w:r></w:p></w:footnote></w:footnotes>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

test('edits footnote text and the paragraph holding its reference', async ({ page }) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'notes.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await footnoteDocx(),
	});
	await expect(fileNameLabel(page)).toHaveText('notes.docx');
	const editor = page.locator('docx-editor');
	const note = editor.locator('.dve-notes li[data-docx-note-id="1"]');
	await expect(note).toContainText('Source');
	await note.dblclick();
	await expect(note.locator('.dve-header-footer-editor .ProseMirror')).toBeFocused();
	await page.keyboard.press('End');
	await page.keyboard.type(': annual report');
	await page.keyboard.press('Escape');
	await expect(note).toContainText('Source: annual report');

	const body = editor.locator('.dve-paper .ProseMirror p').first();
	await body.click();
	await page.keyboard.press('Home');
	await page.keyboard.type('Strong ');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const notes = await zip.file('word/footnotes.xml')!.async('string');
	expect(notes).toContain('Source: annual report');
	expect(notes).toContain('<w:footnoteRef/>');
	const documentXml = await zip.file('word/document.xml')!.async('string');
	expect(documentXml).toContain('Strong Claim');
	expect(documentXml).toContain('<w:footnoteReference w:id="1"/>');
	expect(errors).toEqual([]);
});

test('inserts footnotes in a new document, numbers them in order and saves them', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.dve-paper');
	await surface.locator('p').first().click();
	await page.keyboard.type('Second claim');
	// Ctrl+Alt+F is also bound, but Windows reports Ctrl+Alt as AltGr, so drive the ribbon.
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await editor.getByRole('button', { name: 'Insert footnote', exact: true }).click();
	const secondNote = editor.locator('.dve-notes-footnote li').first();
	await expect(secondNote.locator('.ProseMirror')).toBeFocused();
	await page.keyboard.type('Later source');
	await page.keyboard.press('Escape');

	await surface.locator('p').first().click();
	await page.keyboard.press('Home');
	await page.keyboard.type('First claim');
	await editor.getByRole('button', { name: 'Insert footnote', exact: true }).click();
	await page.keyboard.type('Earlier source');
	await page.keyboard.press('Escape');

	const references = surface.locator('sup.dve-note-reference');
	await expect(references).toHaveText(['1', '2']);
	const notes = editor.locator('.dve-notes-footnote li');
	await expect(notes).toHaveCount(2);
	await expect(notes.first()).toContainText('Earlier source');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const footnotes = await zip.file('word/footnotes.xml')!.async('string');
	expect(footnotes).toContain('Earlier source');
	expect(footnotes).toContain('Later source');
	expect(footnotes).toContain('w:type="separator"');
	const documentXml = await zip.file('word/document.xml')!.async('string');
	expect(documentXml).toMatch(/First claim.*footnoteReference.*Second claim.*footnoteReference/);
	// Text typed next to a reference does not inherit its superscript.
	expect(documentXml).toContain('<w:r><w:t>First claim</w:t></w:r>');
	expect(errors).toEqual([]);
});

test('formats text and inserts a picture inside a header from the ribbon', async ({ page }) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	const png = Buffer.from(
		await page.evaluate(() => {
			const canvas = document.createElement('canvas');
			canvas.width = 24;
			canvas.height = 12;
			canvas.getContext('2d')!.fillRect(0, 0, 24, 12);
			return canvas.toDataURL('image/png').split(',')[1];
		}),
		'base64',
	);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'header.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await headerDocx(),
	});
	const editor = page.locator('docx-editor');
	const header = editor.locator('.dve-header .dve-header-footer-slot');
	await header.dblclick();
	const headerEditor = header.locator('.dve-header-footer-editor .ProseMirror');
	await expect(headerEditor).toBeFocused();
	await page.keyboard.press('Control+a');
	await editor.getByRole('button', { name: 'Bold', exact: true }).click();
	await expect(headerEditor.locator('strong')).toHaveText('Company');
	await page.keyboard.press('End');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const chooser = page.waitForEvent('filechooser');
	await editor.getByRole('button', { name: 'Insert picture', exact: true }).click();
	await (await chooser).setFiles({ name: 'logo.png', mimeType: 'image/png', buffer: png });
	await expect(headerEditor.locator('img[data-docx-image]')).toHaveAttribute('src', /^blob:/);
	await headerEditor.press('Escape');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const headerXml = await zip.file('word/header1.xml')!.async('string');
	expect(headerXml).toContain('<w:b/>');
	expect(headerXml).toContain('<w:drawing>');
	const rels = await zip.file('word/_rels/header1.xml.rels')!.async('string');
	const target = /Target="(media\/dve-picture-[^"]+\.png)"/.exec(rels)?.[1];
	expect(target).toBeDefined();
	expect(await zip.file(`word/${target}`)!.async('nodebuffer')).toEqual(png);
	expect(errors).toEqual([]);
});
