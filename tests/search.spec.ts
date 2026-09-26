import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte']) {
	test(`${framework}: shared find/replace respects formatting, history and readonly`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await expect(surface).toContainText('Document title');
		await page.locator('#new').click();
		await surface.click();
		await page.keyboard.type('alpha ');
		await page.keyboard.press('Control+b');
		await page.keyboard.type('ALPHA');
		await page.keyboard.press('Control+b');
		await page.keyboard.insertText(' café 👩🏽‍💻');
		await page.keyboard.press('Control+f');
		await expect(editor.getByLabel('Find text', { exact: true })).toBeFocused();
		await editor.getByLabel('Find text', { exact: true }).fill('alpha');
		await editor.getByRole('button', { name: 'Find next', exact: true }).click();
		await expect(editor.getByRole('status', { name: 'Search results' })).toContainText('2');
		await editor.getByLabel('Replace with', { exact: true }).fill('$&');
		await editor.getByRole('button', { name: 'Replace all', exact: true }).click();
		await expect(surface).toHaveText('$& $& café 👩🏽‍💻');
		await expect(surface.locator('strong')).toHaveText('$&');
		await editor.getByRole('button', { name: 'Undo', exact: true }).click();
		await expect(surface).toHaveText('alpha ALPHA café 👩🏽‍💻');
		await editor.getByRole('button', { name: 'Redo', exact: true }).click();
		await expect(surface).toHaveText('$& $& café 👩🏽‍💻');
		await page.getByLabel('Read only', { exact: true }).check();
		await editor.getByLabel('Find text', { exact: true }).fill('café');
		await expect(editor.getByRole('button', { name: 'Replace all', exact: true })).toBeDisabled();
		await editor.getByRole('button', { name: 'Find next', exact: true }).click();
		await expect(surface).toHaveText('$& $& café 👩🏽‍💻');
		await editor.getByLabel('Find text', { exact: true }).press('Escape');
		await expect(editor.getByLabel('Find text', { exact: true })).toBeHidden();
		const pending = page.waitForEvent('download');
		await page.locator('#save').click();
		const result = await pending;
		const zip = await JSZip.loadAsync(await readFile((await result.path())!));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('$&amp;');
		expect(xml).toContain('café 👩🏽‍💻');
		expect(errors).toEqual([]);
	});
}
