import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, saveButton } from './helpers';

test('Layout > Page Borders draws a box in Print Layout, undoes and exports pgBorders', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 720 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Bordered page');
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await editor.getByRole('button', { name: 'Page borders', exact: true }).click();
	const dialog = editor.getByRole('dialog', { name: 'Page borders', exact: true });
	await expect(dialog).toBeVisible();
	await dialog.getByRole('combobox', { name: 'Setting', exact: true }).selectOption('box');
	await dialog.getByRole('combobox', { name: 'Style', exact: true }).selectOption('double');
	await dialog.getByRole('combobox', { name: 'Width', exact: true }).selectOption('12');
	// The whole dialog, including its buttons, fits a 720 px high window.
	expect(
		await dialog
			.getByRole('button', { name: 'OK', exact: true })
			.evaluate((button) => button.getBoundingClientRect().bottom <= window.innerHeight),
	).toBe(true);
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(dialog).toBeHidden();
	await editor.locator('#dve-tab-view').click();
	await editor.getByRole('button', { name: 'Print Layout' }).click();
	const border = editor.locator('.dve-print-page .dve-print-page-border');
	await expect(border).toHaveCount(1);
	await expect(border).toHaveCSS('border-top-style', 'double');
	await editor.getByRole('button', { name: 'Web Layout' }).click();
	await editor.locator('.ProseMirror').focus();
	await page.keyboard.press('Control+z');
	await editor.getByRole('button', { name: 'Print Layout' }).click();
	await expect(editor.locator('.dve-print-page')).toHaveCount(1);
	await expect(border).toHaveCount(0);
	await editor.getByRole('button', { name: 'Web Layout' }).click();
	await editor.locator('.ProseMirror').focus();
	await page.keyboard.press('Control+y');
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await downloadPromise).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toMatch(/<w:pgBorders[^>]*>[\s\S]*<w:top w:val="double" w:sz="12"/);
});
