import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { fileInput, insertTableOfSize, newDocument, reveal, saveButton } from './helpers';

for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla']) {
	test(`${framework}: cell shading edits, undoes and survives DOCX reload`, async ({ page }) => {
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		await insertTableOfSize(page);
		const editor = page.locator('docx-editor');
		const cells = editor.locator('.ProseMirror td');
		const open = async () => {
			await editor.getByRole('button', { name: 'Web Layout', exact: true }).click();
			await cells.first().click();
			await editor.getByRole('tab', { name: 'Home', exact: true }).click();
			const menu = editor.getByRole('combobox', { name: 'Borders', exact: true });
			await reveal(editor, menu);
			await menu.focus();
			await menu.press('Enter');
			await editor.getByRole('menuitem', { name: 'Borders and Shading…' }).click();
			return editor.getByRole('dialog', { name: 'Borders and Shading', exact: true });
		};
		let dialog = await open();
		await dialog.getByRole('checkbox', { name: 'No Color', exact: true }).uncheck();
		await expect(dialog.getByLabel('Fill', { exact: true })).toBeEnabled();
		await dialog.getByLabel('Fill', { exact: true }).fill('#123456');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(cells.first()).toHaveCSS('background-color', 'rgb(18, 52, 86)');
		await page.keyboard.press('Control+z');
		await expect(cells.first()).not.toHaveCSS('background-color', 'rgb(18, 52, 86)');
		await page.keyboard.press('Control+y');
		await expect(cells.first()).toHaveCSS('background-color', 'rgb(18, 52, 86)');
		const save = async () => {
			const pending = page.waitForEvent('download');
			await saveButton(page).click();
			return readFile((await (await pending).path())!);
		};
		const bytes = await save();
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('w:fill="123456"');
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'shading.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: bytes,
		});
		await expect(cells.first()).toHaveCSS('background-color', 'rgb(18, 52, 86)');
		dialog = await open();
		await dialog.getByRole('combobox', { name: 'Apply to', exact: true }).selectOption('table');
		await dialog.getByRole('checkbox', { name: 'No Color', exact: true }).check();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		for (let i = 0; i < 4; i++)
			await expect(cells.nth(i)).not.toHaveCSS('background-color', 'rgb(18, 52, 86)');
		const cleared = await save();
		const clearedXml = await (
			await JSZip.loadAsync(cleared)
		)
			.file('word/document.xml')!
			.async('string');
		expect(clearedXml.match(/w:fill="auto"/g)).toHaveLength(4);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'cleared.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: cleared,
		});
		await expect(cells.first()).not.toHaveCSS('background-color', 'rgb(18, 52, 86)');
	});
}
