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
	test(`${framework}: line breaks and spacing survive editing and DOCX reload`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await openSample(page, framework);
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await expect(surface).toContainText('Document title');
		await newDocument(page);
		await surface.click();
		await page.keyboard.type('First line');
		await page.keyboard.press('Shift+Enter');
		await page.keyboard.type('Second line');
		await expect(surface.locator('p')).toHaveCount(1);
		await expect(surface.locator('p br:not(.ProseMirror-trailingBreak)')).toHaveCount(1);
		await page.keyboard.press('Enter');
		await page.keyboard.type('Next paragraph');
		await expect(surface.locator('p')).toHaveCount(2);
		await page.keyboard.press('Control+a');
		const spacing = editor.getByLabel('Line spacing', { exact: true });
		await spacing.selectOption('auto:360');
		await spacing.selectOption('auto:480');
		await editor.getByRole('button', { name: 'Undo', exact: true }).click();
		await expect(spacing).toHaveValue('auto:360');
		await editor.getByRole('button', { name: 'Redo', exact: true }).click();
		await expect(spacing).toHaveValue('auto:480');
		for (const paragraph of await surface.locator('p').all())
			await expect(paragraph).toHaveAttribute('style', /line-height:\s*2\.4(?:;|$)/);
		await setReadOnly(page, true);
		await expect(spacing).toBeDisabled();
		await setReadOnly(page, false);

		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const downloaded = await pending;
		const buffer = await readFile((await downloaded.path())!);
		const zip = await JSZip.loadAsync(buffer);
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml.match(/<w:br\s*\/>/g)).toHaveLength(1);
		expect(xml.match(/w:line="480"/g)).toHaveLength(2);
		expect(xml.match(/w:lineRule="auto"/g)).toHaveLength(2);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'paragraph-roundtrip.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer,
		});
		await expect(fileNameLabel(page)).toHaveText('paragraph-roundtrip.docx');
		await expect(surface.locator('p')).toHaveCount(2);
		await expect(surface.locator('p br:not(.ProseMirror-trailingBreak)')).toHaveCount(1);
		await surface.locator('p').first().click();
		await expect(spacing).toHaveValue('auto:480');
		await spacing.selectOption('inherit');
		await expect(spacing).toHaveValue('inherit');
		await page.keyboard.press('Control+a');
		await expect(spacing).toHaveValue('mixed');
		expect(errors).toEqual([]);
	});
}
