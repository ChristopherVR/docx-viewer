import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, saveButton } from './helpers';

test('Layout > Watermark prints behind the text in Print Layout, undoes and exports VML', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 720 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Body text');
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await editor.getByRole('button', { name: 'Watermark', exact: true }).click();
	const dialog = editor.getByRole('dialog', { name: 'Printed watermark', exact: true });
	await expect(dialog).toBeVisible();
	await dialog.getByRole('combobox', { name: 'Setting', exact: true }).selectOption('text');
	await dialog.getByRole('combobox', { name: 'Text', exact: true }).fill('CONFIDENTIAL');
	expect(
		await dialog
			.getByRole('button', { name: 'OK', exact: true })
			.evaluate((button) => button.getBoundingClientRect().bottom <= window.innerHeight),
	).toBe(true);
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(dialog).toBeHidden();
	await editor.locator('#dve-tab-view').click();
	await editor.getByRole('button', { name: 'Print Layout' }).click();
	const mark = editor.locator('.dve-print-page .dve-print-watermark');
	await expect(mark).toHaveCount(1);
	await expect(mark).toHaveText('CONFIDENTIAL');
	await expect(mark).toHaveCSS('opacity', '0.5');
	await editor.getByRole('button', { name: 'Web Layout' }).click();
	await editor.locator('.ProseMirror').focus();
	await page.keyboard.press('Control+z');
	await editor.getByRole('button', { name: 'Print Layout' }).click();
	await expect(editor.locator('.dve-print-page')).toHaveCount(1);
	await expect(mark).toHaveCount(0);
	await editor.getByRole('button', { name: 'Web Layout' }).click();
	await editor.locator('.ProseMirror').focus();
	await page.keyboard.press('Control+y');
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await downloadPromise;
	await download.saveAs('test-results/watermark.docx');
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const header = Object.keys(zip.files).find((name) => /^word\/header\d*\.xml$/.test(name));
	expect(header).toBeTruthy();
	const xml = await zip.file(header!)!.async('string');
	expect(xml).toMatch(/id="PowerPlusWaterMarkObject\d+"/);
	expect(xml).toContain('string="CONFIDENTIAL"');
	expect(xml).toContain('rotation:315');
});
