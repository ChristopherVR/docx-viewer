import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { insertTableOfSize, newDocument, saveButton, setReadOnly } from './helpers';

test('Table Properties edits, undoes, validates and exports Word row settings', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	const cell = editor.locator('.ProseMirror td').first();
	await cell.click();
	await editor.getByRole('tab', { name: 'Table', exact: true }).click();
	const properties = editor.getByRole('button', { name: 'Table properties', exact: true });
	await properties.click();
	const dialog = editor.getByRole('dialog', { name: 'Table properties', exact: true });
	await dialog.getByRole('checkbox', { name: 'Specify height', exact: true }).check();
	await dialog.getByRole('spinbutton', { name: 'Height (inches)', exact: true }).fill('0.4');
	await dialog.getByRole('combobox', { name: 'Row height is', exact: true }).selectOption('exact');
	await dialog
		.getByRole('checkbox', { name: 'Allow row to break across pages', exact: true })
		.uncheck();
	await dialog
		.getByRole('checkbox', { name: 'Repeat as header row at the top of each page', exact: true })
		.check();
	await dialog.getByRole('spinbutton', { name: 'Top', exact: true }).fill('-1');
	await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled();
	await dialog.getByRole('spinbutton', { name: 'Top', exact: true }).fill('0.1');
	await page.screenshot({ path: 'test-results/table-properties.png' });
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(dialog).toBeHidden();
	await expect(editor.locator('.ProseMirror')).toBeFocused();
	await expect(editor.locator('.ProseMirror tr').first()).toHaveAttribute(
		'style',
		/height:\s*38\.4px/,
	);
	await page.keyboard.press('Control+z');
	await expect(editor.locator('.ProseMirror tr').first()).not.toHaveAttribute('style');
	await page.keyboard.press('Control+y');
	await properties.click();
	await expect(
		dialog.getByRole('spinbutton', { name: 'Height (inches)', exact: true }),
	).toHaveValue('0.4');
	await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await downloadPromise;
	await download.saveAs('test-results/table-properties.docx');
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toContain('<w:cantSplit');
	expect(xml).toContain('w:val="576" w:hRule="exact"');
	expect(xml).toContain('<w:tblHeader');
	expect(xml).toContain('<w:top w:w="144"');
	await setReadOnly(page, true);
	await expect(properties).toBeDisabled();
});

test('Table Properties stays usable on a narrow ribbon and localizes the dialog', async ({
	page,
}) => {
	await page.setViewportSize({ width: 720, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror td').first().click();
	await editor.getByRole('tab', { name: 'Table', exact: true }).click();
	await editor.getByRole('button', { name: 'Table properties', exact: true }).click();
	await expect(editor.getByRole('dialog', { name: 'Table properties' })).toBeVisible();
	await page.keyboard.press('Escape');
	await editor.evaluate((element) => {
		element.setAttribute('locale', 'de');
	});
	await editor.getByRole('button', { name: 'Tabelleneigenschaften', exact: true }).click();
	await expect(editor.getByRole('dialog', { name: 'Tabelleneigenschaften' })).toBeVisible();
	await expect(editor.getByRole('checkbox', { name: 'Höhe angeben' })).toBeVisible();
});
