import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, reveal, saveButton } from './helpers';

test('More Columns applies spacing and separator with undo and export', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	await page.keyboard.type('A column layout');
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const columns = editor.getByRole('combobox', { name: 'Columns', exact: true });
	await reveal(editor, columns);
	await columns.press('Alt+ArrowDown');
	await editor.getByRole('menuitem', { name: 'More Columns…', exact: true }).click();
	const dialog = editor.getByRole('dialog', { name: 'Columns', exact: true });
	await dialog.getByRole('spinbutton', { name: 'Number of columns', exact: true }).fill('3');
	await dialog.getByRole('spinbutton', { name: 'Column spacing (inches)', exact: true }).fill('5');
	await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled();
	await dialog
		.getByRole('spinbutton', { name: 'Column spacing (inches)', exact: true })
		.fill('0.25');
	await dialog.getByRole('checkbox', { name: 'Line between', exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(columns).toHaveValue('3');
	await expect(editor.locator('.dve-paper')).toHaveCSS('column-gap', '24px');
	await page.keyboard.press('Control+z');
	await expect(columns).toHaveValue('1');
	await page.keyboard.press('Control+y');
	await expect(columns).toHaveValue('3');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
	await expect(editor.locator('.dve-print-column-rule')).toHaveCount(2);
	const promise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await promise;
	await download.saveAs('test-results/columns-options.docx');
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toMatch(/<w:cols[^>]*w:num="3"[^>]*w:space="360"[^>]*w:sep="1"/);
});

test('Left and Right presets render unequal widths and More Columns edits them', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	// Typing ~6,000 characters key by key exceeds the test timeout on slow CI runners.
	await page.keyboard.insertText(
		'A paragraph with words that wrap through unequal columns. '.repeat(100),
	);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const columns = editor.getByRole('combobox', { name: 'Columns', exact: true });
	await reveal(editor, columns);
	await columns.selectOption('left');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	const view = editor.getByRole('combobox', { name: 'Layout view', exact: true });
	await view.selectOption('print');
	await expect(editor.locator('.dve-print-column').nth(0)).toHaveCSS('width', '192px');
	await expect(editor.locator('.dve-print-column').nth(1)).toHaveCSS('width', '384px');
	await view.selectOption('draft');
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await reveal(editor, columns);
	await columns.selectOption('right');
	await expect(columns).toHaveValue('right');
	await columns.press('Alt+ArrowDown');
	await editor.getByRole('menuitem', { name: 'More Columns…', exact: true }).click();
	const dialog = editor.getByRole('dialog', { name: 'Columns', exact: true });
	await expect(
		dialog.getByRole('checkbox', { name: 'Equal column width', exact: true }),
	).not.toBeChecked();
	const first = dialog
		.getByRole('group', { name: 'Column 1', exact: true })
		.getByRole('spinbutton', { name: 'Column width (inches)', exact: true });
	await first.fill('9');
	await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled();
	await first.fill('1.5');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await page.keyboard.press('Control+z');
	await expect(columns).toHaveValue('right');
	await page.keyboard.press('Control+y');
	await expect(columns).toHaveValue('');
	const promise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await promise;
	await download.saveAs('test-results/unequal-columns.docx');
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toContain('w:equalWidth="0"');
	expect(xml).toContain('w:w="2160"');
	expect(xml).toContain('w:w="6480"');
});
