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
			const data = canvas.toDataURL('image/png').split(',')[1];
			if (!data) throw new Error('The canvas produced no PNG data.');
			return data;
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
		(name) => name.startsWith('word/media/') && !zip.files[name]?.dir,
	);
	expect(media).toHaveLength(1);
	const [mediaName] = media;
	if (!mediaName) throw new Error('The saved document has no media part.');
	expect(await zip.file(mediaName)!.async('nodebuffer')).toEqual(PNG);
	expect(errors).toEqual([]);
});

test('inserts an SVG picture with a PNG fallback', async ({ page }) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const chooser = page.waitForEvent('filechooser');
	await editor.getByRole('button', { name: 'Insert picture', exact: true }).click();
	const svg =
		'<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><rect width="40" height="20" fill="#185abd"/></svg>';
	await (
		await chooser
	).setFiles({
		name: 'badge.svg',
		mimeType: 'image/svg+xml',
		buffer: Buffer.from(svg),
	});
	const picture = surface.locator('img[data-docx-image]');
	await expect(picture).toHaveAttribute('width', '40');
	const shown = await picture.evaluate(async (img: HTMLImageElement) =>
		(await fetch(img.src)).headers.get('content-type'),
	);
	expect(shown).toBe('image/svg+xml');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const media = Object.keys(zip.files).filter(
		(name) => name.startsWith('word/media/') && !zip.files[name]?.dir,
	);
	expect(media.map((name) => name.split('.').pop()).sort()).toEqual(['png', 'svg']);
	const png = await zip.file(media.find((name) => name.endsWith('.png'))!)!.async('uint8array');
	expect([...png.slice(0, 4)]).toEqual([137, 80, 78, 71]);
	expect(await zip.file('word/document.xml')!.async('string')).toContain('asvg:svgBlip');
	expect(errors).toEqual([]);
});
