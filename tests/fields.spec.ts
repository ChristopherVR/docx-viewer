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
