import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { insertTableOfSize, newDocument, reveal, saveButton } from './helpers';

test('Home > Borders draws cell edges in a table, undoes, and exports tcBorders', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	const first = editor.locator('.ProseMirror td').first();
	const below = editor.locator('.ProseMirror tr').nth(1).locator('td').first();
	const edge = (cell: typeof first, side: 'Top' | 'Bottom' | 'Left') =>
		cell.evaluate(
			(element, name) => getComputedStyle(element)[`border${name}Style` as 'borderTopStyle'],
			side,
		);
	await first.click();
	await editor.getByRole('tab', { name: 'Home', exact: true }).click();
	const menu = editor.getByRole('combobox', { name: 'Borders' });
	await reveal(editor, menu);
	// Inserted tables start with a grid, so clear the cell's edges first.
	await menu.selectOption('none');
	await expect.poll(() => edge(first, 'Bottom')).not.toBe('solid');
	await menu.selectOption('bottom');
	await expect.poll(() => edge(first, 'Bottom')).toBe('solid');
	// The cell below shares the edge, so it draws it as its top border.
	expect(await edge(below, 'Top')).toBe('solid');
	await editor.locator('.ProseMirror').focus();
	await page.keyboard.press('Control+z');
	await expect.poll(() => edge(first, 'Bottom')).not.toBe('solid');
	await page.keyboard.press('Control+y');
	await expect.poll(() => edge(first, 'Bottom')).toBe('solid');
	await first.click();
	await menu.selectOption('left');
	await expect.poll(() => edge(first, 'Left')).toBe('solid');
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await downloadPromise).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toMatch(/<w:tcBorders>[\s\S]*?<w:bottom /);
	expect(xml).toMatch(/<w:tcBorders>[\s\S]*?<w:left /);
});
