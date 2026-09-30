import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type { EditorView } from 'prosemirror-view';
import type { DocumentModel } from '@christophervr/docx-core';
import { newDocument, reveal, saveButton } from './helpers';

test('Line Numbering Options, paragraph suppression, undo and DOCX export agree', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	await page.keyboard.type('First paragraph');
	await page.keyboard.press('Enter');
	await page.keyboard.type('Second paragraph');
	await page.keyboard.press('Enter');
	await page.keyboard.type('Third paragraph');
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const menu = editor.getByRole('combobox', { name: 'Line numbers', exact: true });
	await reveal(editor, menu);
	await menu.selectOption('options');
	const dialog = editor.getByRole('dialog', { name: 'Line numbers', exact: true });
	await dialog.getByRole('checkbox', { name: 'Add line numbering', exact: true }).check();
	await dialog.getByRole('spinbutton', { name: 'Start at', exact: true }).fill('3');
	await dialog.getByRole('spinbutton', { name: 'Count by', exact: true }).fill('0');
	await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled();
	await dialog.getByRole('spinbutton', { name: 'Count by', exact: true }).fill('1');
	await dialog
		.getByRole('checkbox', { name: 'Automatic distance from text', exact: true })
		.uncheck();
	await dialog
		.getByRole('spinbutton', { name: 'Distance from text (inches)', exact: true })
		.fill('0.25');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(dialog).toBeHidden();
	await expect(menu).toHaveValue('continuous');
	await page.keyboard.press('Control+z');
	await expect(menu).toHaveValue('none');
	await expect(surface).toContainText('First paragraph');
	await page.keyboard.press('Control+y');
	await expect(menu).toHaveValue('continuous');
	await surface
		.locator('p')
		.nth(1)
		.click({ position: { x: 12, y: 8 } });
	await expect
		.poll(() =>
			editor.evaluate(
				(element) =>
					(element as unknown as { view: EditorView }).view.state.selection.$from.parent
						.textContent,
			),
		)
		.toBe('Second paragraph');
	await menu.selectOption('suppress');
	await expect
		.poll(() =>
			editor.evaluate((element) =>
				(element as unknown as { documentModel: DocumentModel }).documentModel.blocks.find(
					(block) =>
						block.type === 'paragraph' &&
						block.runs.map((run) => run.text).join('') === 'Second paragraph',
				),
			),
		)
		.toMatchObject({ suppressLineNumbers: true });
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
	await expect(editor.locator('.dve-print-line-number')).toHaveText(['3', '4']);
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await downloadPromise;
	await download.saveAs('test-results/line-number-options.docx');
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toMatch(/<w:lnNumType[^>]*w:start="2"[^>]*w:distance="360"/);
	expect(xml).toContain('<w:suppressLineNumbers');
	await page.screenshot({ path: 'test-results/line-number-options.png' });
});
