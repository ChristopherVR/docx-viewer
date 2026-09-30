import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, reveal, saveButton } from './helpers';

test('Drop Cap Options apply, undo, cancel and export the frame font and distance', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	await page.keyboard.type('Once upon a time');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const menu = editor.getByRole('combobox', { name: 'Drop cap', exact: true });
	await reveal(editor, menu);
	await menu.selectOption('options');
	const dialog = editor.getByRole('dialog', { name: 'Drop Cap Options', exact: true });
	await dialog.getByRole('combobox', { name: 'Position', exact: true }).selectOption('drop');
	await dialog.getByRole('textbox', { name: 'Font family', exact: true }).fill('Georgia');
	await dialog.getByRole('spinbutton', { name: 'Lines to drop', exact: true }).fill('1.5');
	await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled();
	await dialog.getByRole('spinbutton', { name: 'Lines to drop', exact: true }).fill('4');
	await dialog
		.getByRole('spinbutton', { name: 'Distance from text (inches)', exact: true })
		.fill('0.1');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(surface.locator('p')).toHaveCount(2);
	await expect(surface.locator('p').first()).toHaveText('O');
	await expect(surface.locator('p').first()).toHaveCSS('margin-right', '9.6px');
	await page.keyboard.press('Control+z');
	await expect(surface.locator('p')).toHaveCount(1);
	await expect(surface).toHaveText('Once upon a time');
	await page.keyboard.press('Control+y');
	await expect(surface.locator('p')).toHaveCount(2);
	await reveal(editor, menu);
	await menu.selectOption('options');
	await expect(dialog.getByRole('spinbutton', { name: 'Lines to drop', exact: true })).toHaveValue(
		'4',
	);
	await expect(dialog.getByRole('textbox', { name: 'Font family', exact: true })).toHaveValue(
		'Georgia',
	);
	await dialog.getByRole('spinbutton', { name: 'Lines to drop', exact: true }).fill('2');
	await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await downloadPromise;
	await download.saveAs('test-results/drop-cap-options.docx');
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toContain('w:dropCap="drop"');
	expect(xml).toContain('w:lines="4"');
	expect(xml).toContain('w:hSpace="144"');
	expect(xml).toContain('w:ascii="Georgia"');
	await page.screenshot({ path: 'test-results/drop-cap-options.png' });
});
