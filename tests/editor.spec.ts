import { test, expect } from '@playwright/test';
import {
	openSample,
	newDocument,
	fileInput,
	saveButton,
	saveCopyAsDocx,
	fileNameLabel,
	saveStateLabel,
	setReadOnly,
	insertTableOfSize,
} from './helpers';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: shared editor edits, formats, saves, reloads, and respects read-only`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await openSample(page, framework);
		await expect(page.getByRole('navigation', { name: 'Choose framework' })).toHaveCount(0);
		const editor = page.locator('docx-editor');
		const text = editor.locator('.ProseMirror');
		await expect(text).toContainText('Document title');
		await text.locator('p').last().click();
		await page.keyboard.press('End');
		await page.keyboard.type(' Browser contract.');
		await expect(saveStateLabel(page)).toHaveText('Unsaved changes');
		await editor.getByRole('button', { name: 'Bold', exact: true }).click();
		await page.keyboard.type(' Bold contract.');
		await expect(text.locator('strong').filter({ hasText: 'Bold contract.' })).toHaveCount(1);
		await editor.getByRole('button', { name: 'Undo', exact: true }).click();
		await expect(text).not.toContainText('Bold contract.');
		await editor.getByRole('button', { name: 'Redo', exact: true }).click();
		await expect(text).toContainText('Bold contract.');
		await setReadOnly(page, true);
		await expect(text).toHaveAttribute('contenteditable', 'false');
		await expect(editor.getByRole('button', { name: 'Bold', exact: true })).toBeDisabled();
		await setReadOnly(page, false);
		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const download = await pending;
		const path = await download.path();
		const zip = await JSZip.loadAsync(await readFile(path!));
		expect(await zip.file('word/document.xml')!.async('string')).toContain('Bold contract.');
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'reopened.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await readFile(path!),
		});
		await expect(fileNameLabel(page)).toHaveText('reopened.docx');
		await expect(text).toContainText('Browser contract.');
		await expect(text).toContainText('Bold contract.');
		expect(errors).toEqual([]);
	});
}
test('imports and saves legacy DOC as DOC and exports visible text as DOCX', async ({ page }) => {
	await page.goto('/');
	await (
		await fileInput(page)
	).setInputFiles('packages/legacy/src/__tests__/fixtures/ole-word-97.doc');
	await expect(fileNameLabel(page)).toHaveText('ole-word-97.doc');
	const original = await readFile('packages/legacy/src/__tests__/fixtures/ole-word-97.doc');
	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await pending;
	expect(await readFile((await download.path())!)).toEqual(original);
	const editable = page.locator('docx-editor .ProseMirror');
	await editable.locator('p').first().click();
	await page.keyboard.press('End');
	await page.keyboard.type(' Legacy edit.');
	const editedDownload = page.waitForEvent('download');
	await saveButton(page).click();
	const editedFile = await editedDownload;
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'edited.doc',
		mimeType: 'application/msword',
		buffer: await readFile((await editedFile.path())!),
	});
	await expect(fileNameLabel(page)).toHaveText('edited.doc');
	await expect(editable).toContainText('Legacy edit.');
	const exported = page.waitForEvent('download');
	await saveCopyAsDocx(page);
	const docx = await exported;
	expect(docx.suggestedFilename()).toBe('edited.docx');
	const zip = await JSZip.loadAsync(await readFile((await docx.path())!));
	expect(zip.file('word/document.xml')).not.toBeNull();
});

test('ribbon edits preserve font properties and save table and page settings', async ({ page }) => {
	await page.goto('/');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	await page.keyboard.type('Ribbon document');
	await page.keyboard.press('Control+a');
	await editor.getByLabel('Font family', { exact: true }).fill('Georgia');
	await page.keyboard.press('Enter');
	await editor.getByLabel('Font size', { exact: true }).fill('18');
	await page.keyboard.press('Enter');
	await editor.getByRole('button', { name: 'Font color options' }).click();
	await editor.getByRole('menuitem', { name: 'Blue', exact: true }).click();
	await expect(surface.locator('span').filter({ hasText: 'Ribbon document' }).last()).toHaveCSS(
		'font-family',
		'Georgia',
	);
	await expect(surface.locator('span').filter({ hasText: 'Ribbon document' }).last()).toHaveCSS(
		'font-size',
		'24px',
	);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await editor.getByLabel('Margins', { exact: true }).selectOption('narrow');
	await editor.getByLabel('Orientation', { exact: true }).selectOption('landscape');
	await expect(editor.getByLabel('Document page', { exact: true })).toHaveCSS('width', '1056px');
	await editor.getByRole('tab', { name: 'Home', exact: true }).click();
	await editor.getByRole('button', { name: 'Undo', exact: true }).click();
	await expect(editor.getByLabel('Document page', { exact: true })).toHaveCSS('width', '816px');
	await editor.getByRole('button', { name: 'Redo', exact: true }).click();
	await surface.click();
	await page.keyboard.press('Control+End');
	await page.keyboard.press('Enter');
	await insertTableOfSize(page);
	await expect(surface.locator('table')).toHaveCount(1);
	await setReadOnly(page, true);
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByLabel('Zoom', { exact: true }).selectOption('75');
	await expect(surface).toHaveAttribute('contenteditable', 'false');
	await expect(editor.getByLabel('Document page', { exact: true })).toHaveCSS('zoom', '0.75');
	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const result = await pending;
	const zip = await JSZip.loadAsync(await readFile((await result.path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toContain('Georgia');
	expect(xml).toContain('0070c0');
	expect(xml).toContain('w:val="36"');
	expect(xml).toContain('<w:tbl>');
	expect(xml).toContain('w:w="15840"');
});
