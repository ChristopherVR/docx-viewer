import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, saveButton } from './helpers';

test('inserts a picture and a hyperlink from the ribbon and saves both to DOCX', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	const PNG = Buffer.from(
		await page.evaluate(() => {
			const canvas = document.createElement('canvas');
			canvas.width = 40;
			canvas.height = 20;
			canvas.getContext('2d')!.fillRect(0, 0, 40, 20);
			return canvas.toDataURL('image/png').split(',')[1];
		}),
		'base64',
	);
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	await page.keyboard.type('Read the guide');

	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const chooser = page.waitForEvent('filechooser');
	await editor.getByRole('button', { name: 'Insert picture', exact: true }).click();
	await (await chooser).setFiles({ name: 'dot.png', mimeType: 'image/png', buffer: PNG });
	const picture = surface.locator('img[data-docx-image]');
	await expect(picture).toHaveAttribute('src', /^blob:/);
	await expect(picture).toHaveAttribute('alt', 'dot');

	await surface.locator('p').first().click();
	await page.keyboard.press('Home');
	await page.keyboard.press('Shift+End');
	await page.keyboard.press('Control+k');
	const dialog = editor.getByRole('dialog', { name: 'Insert link' });
	await expect(dialog).toBeVisible();
	await dialog.getByLabel('Address', { exact: true }).fill('https://example.com/guide');
	await dialog.getByRole('button', { name: 'Insert', exact: true }).click();
	await expect(surface.locator('a[data-docx-link]')).toHaveAttribute(
		'href',
		'https://example.com/guide',
	);

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const documentXml = await zip.file('word/document.xml')!.async('string');
	const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
	expect(documentXml).toContain('<w:hyperlink');
	expect(documentXml).toContain('<w:drawing>');
	expect(rels).toContain('Target="https://example.com/guide"');
	expect(rels).toMatch(/Target="media\/dve-picture-[^"]+\.png"/);
	const media = Object.keys(zip.files).filter(
		(name) => name.startsWith('word/media/') && !zip.files[name].dir,
	);
	expect(media).toHaveLength(1);
	expect(await zip.file(media[0])!.async('nodebuffer')).toEqual(PNG);
	expect(errors).toEqual([]);
});
