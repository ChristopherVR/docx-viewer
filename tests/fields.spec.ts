import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { fileInput, fileNameLabel, saveButton } from './helpers';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fieldDocx(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t xml:space="preserve">Author: </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> AUTHOR </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Ann</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> wrote this.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

test('edits text around and inside a complex field and saves the field intact', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'field.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await fieldDocx(),
	});
	await expect(fileNameLabel(page)).toHaveText('field.docx');
	const editor = page.locator('docx-editor');
	const body = editor.locator('.ProseMirror').first();
	await expect(body).toContainText('Author: Ann wrote this.');
	await expect(body).toHaveAttribute('contenteditable', 'true');

	await editor.locator('[data-field="AUTHOR"]').dblclick();
	// Windows word selection includes the trailing space, crossing the hidden field end marker.
	await page.keyboard.type('Bea');
	await page.keyboard.press('End');
	await page.keyboard.type(' Twice.');
	await expect(body).toContainText(/Author: Bea ?wrote this\. Twice\./);

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toMatch(
		/<w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText xml:space="preserve"> AUTHOR <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t>Bea<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/>/,
	);
	expect(xml).toMatch(/ ?wrote this\. Twice\./);
	expect(errors).toEqual([]);
});

async function headingsDocx(): Promise<Buffer> {
	const zip = await JSZip.loadAsync(await fieldDocx());
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body><w:p/><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2"/><w:pageBreakBefore/></w:pPr><w:r><w:t>Details</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

test('inserts a table of contents from the References tab and saves it as a TOC field', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'headings.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await headingsDocx(),
	});
	await expect(fileNameLabel(page)).toHaveText('headings.docx');
	const editor = page.locator('docx-editor');
	const body = editor.locator('.ProseMirror').first();
	await body.locator('p').first().click();
	await editor.getByRole('tab', { name: 'References', exact: true }).click();
	await editor.getByRole('button', { name: 'Insert table of contents', exact: true }).click();
	await expect(body.locator('p').first()).toHaveText(/^Overview\s+1$/);
	await expect(body.locator('p').nth(1)).toHaveText(/^Details\s+2$/);

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toMatch(/<w:instrText xml:space="preserve"> TOC [^<]*<\/w:instrText>/);
	expect(xml).toContain('<w:tab w:val="right" w:leader="dot" w:pos="9360"/>');
	// Entries are hyperlinks to _Toc bookmarks on the headings, with PAGEREF page numbers.
	expect(xml).toMatch(
		/<w:hyperlink w:anchor="(_Toc\d+)"[^>]*><w:r><w:t>Details<\/w:t><w:tab\/><\/w:r>.*?PAGEREF \1 .*?<w:t>2<\/w:t>/,
	);
	expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Toc\d+"\/><w:r><w:t>Details<\/w:t>/);
	expect(errors).toEqual([]);
});
