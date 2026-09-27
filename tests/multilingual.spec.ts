import { expect, test } from '@playwright/test';
import {
	openSample,
	newDocument,
	fileInput,
	saveButton,
	fileNameLabel,
	setReadOnly,
} from './helpers';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: multilingual metadata and RTL survive DOCX save/reload`, async ({ page }) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await openSample(page, framework);
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await expect(surface).toContainText('Document title');
		await newDocument(page);
		await surface.click();
		const text = 'مرحبا بالعالم שלום עולם 日本語 cafe\u0301 👩🏽‍💻';
		await page.keyboard.insertText(text);
		await page.keyboard.press('Control+a');
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		// Multilingual controls are shared by every framework adapter.
		await editor.getByLabel('Paragraph direction', { exact: true }).selectOption('rtl');
		await editor.getByLabel('Text language', { exact: true }).selectOption('ar-SA');
		await editor.getByLabel('Run direction', { exact: true }).selectOption('on');
		await expect(surface.locator('p')).toHaveAttribute('dir', 'rtl');
		await expect(surface.locator('[lang="ar-SA"]').first()).toContainText('مرحبا');
		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const result = await pending;
		const buffer = await readFile((await result.path())!);
		const zip = await JSZip.loadAsync(buffer);
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:bidi');
		expect(xml).toContain('<w:rtl');
		expect(xml).toContain('w:val="ar-SA"');
		expect(xml).toContain(text);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'multilingual.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer,
		});
		await expect(fileNameLabel(page)).toHaveText('multilingual.docx');
		await expect(surface.locator('p')).toHaveAttribute('dir', 'rtl');
		await expect(surface).toHaveText(text);
		await setReadOnly(page, true);
		await expect(editor.getByLabel('Paragraph direction', { exact: true })).toBeDisabled();
		await expect(editor.getByLabel('Text language', { exact: true })).toBeDisabled();
		expect(errors).toEqual([]);
	});
}
