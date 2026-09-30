import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, reveal, saveButton } from './helpers';

test('Layout > Page Color paints the page, saves w:background and undoes', async ({ page }) => {
	await page.setViewportSize({ width: 1700, height: 800 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const paper = editor.getByLabel('Document page', { exact: true });
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Tinted page');
	await editor.locator('#dve-tab-layout').click();
	const caret = editor.locator('[aria-label="Page color"] + [data-split-caret]');
	await reveal(editor, caret);
	await caret.click();
	await editor.getByRole('menuitem', { name: 'Light Green', exact: true }).click();
	await expect(paper).toHaveCSS('background-color', 'rgb(146, 208, 80)');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	expect(await zip.file('word/document.xml')!.async('string')).toContain(
		'<w:background w:color="92D050"/>',
	);
	expect(await zip.file('word/settings.xml')!.async('string')).toContain(
		'<w:displayBackgroundShape/>',
	);

	await editor.locator('.ProseMirror').click();
	await page.keyboard.press('Control+z');
	await expect(paper).toHaveCSS('background-color', 'rgb(255, 255, 255)');
});

test('Layout > Hyphenation turns automatic hyphenation on, saves it and undoes', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 800 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const paper = editor.getByLabel('Document page', { exact: true });
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Hyphenate me');
	await editor.locator('#dve-tab-layout').click();
	const select = editor.getByLabel('Hyphenation', { exact: true });
	await reveal(editor, select);
	await select.selectOption('auto');
	await expect(paper).toHaveCSS('hyphens', 'auto');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	expect(await zip.file('word/settings.xml')!.async('string')).toContain('<w:autoHyphenation/>');

	await editor.locator('.ProseMirror').click();
	await page.keyboard.press('Control+z');
	await expect(paper).toHaveCSS('hyphens', 'manual');
});
