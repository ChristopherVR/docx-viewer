import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { fileInput, fileNameLabel, newDocument, saveButton } from './helpers';

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
	// The editing surface sizes the tab to the right-aligned stop: the page number ends at the
	// right margin, and the tab shows a dot leader.
	await expect(body.locator('.dve-tab-leader-dot')).toHaveCount(2);
	await expect
		.poll(() =>
			body.evaluate((root) => {
				const paragraph = root.querySelectorAll('p')[1];
				if (!paragraph) throw new Error('The second paragraph is missing.');
				const range = document.createRange();
				range.selectNodeContents(paragraph);
				return Math.round(root.getBoundingClientRect().right - range.getBoundingClientRect().right);
			}),
		)
		.toBeLessThanOrEqual(2);

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

test('records cut and paste under Track Changes as a tracked move and saves it', async ({
	page,
	context,
}) => {
	await context.grantPermissions(['clipboard-read', 'clipboard-write']);
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const body = editor.locator('.ProseMirror').first();
	await body.locator('p').first().click();
	await page.keyboard.type('Alpha Beta');
	await page.keyboard.press('Enter');
	await page.keyboard.type('Gamma');
	await editor.getByRole('tab', { name: 'Review', exact: true }).click();
	await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
	// Select "Beta" and cut it, then paste it at the end of "Gamma".
	await body
		.locator('p')
		.first()
		.dblclick({ position: { x: 60, y: 8 } });
	await page.keyboard.press('Control+x');
	await page.keyboard.press('ArrowDown');
	await page.keyboard.press('End');
	// Let the editor read the caret move (a selectionchange task) before pasting.
	await page.evaluate(
		() => new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 50))),
	);
	await page.keyboard.press('Control+v');
	await expect(editor.locator('.dve-revision-move')).toHaveCount(2);

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	const name = /<w:moveFromRangeStart w:id="\d+"[^>]* w:name="(move\d+)"\/>/.exec(xml)?.[1];
	expect(name).toBeTruthy();
	expect(xml).toContain(`<w:moveToRangeStart`);
	expect(xml).toMatch(/<w:moveFrom w:id="\d+"[^>]*><w:r><w:delText>Beta<\/w:delText>/);
	expect(xml).toMatch(/<w:moveTo w:id="\d+"[^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t>Beta<\/w:t>/);
	// Pasted CSS values are converted to what Word accepts, never written verbatim.
	expect(xml).not.toMatch(/rgb\(|sans-serif/);
	expect(xml).not.toContain('dve-rev');
	expect(errors).toEqual([]);
});
