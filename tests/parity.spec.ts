import { expect, test } from '@playwright/test';
import {
	openSample,
	newDocument,
	fileInput,
	saveButton,
	fileNameLabel,
	setReadOnly,
	insertTableOfSize,
} from './helpers';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: Word run formatting and table editing parity`, async ({ page }) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await openSample(page, framework);
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await expect(surface).toContainText('Document title');

		// Apply the Word-only run properties to a selected string through the ribbon.
		await newDocument(page);
		await surface.click();
		await page.keyboard.type('Parity formatting');
		await page.keyboard.press('Control+a');
		await editor.getByRole('button', { name: 'Strikethrough', exact: true }).click();
		await editor.getByRole('button', { name: 'Text highlight options' }).click();
		await editor.getByRole('menuitem', { name: 'Cyan', exact: true }).click();
		await editor.getByRole('button', { name: 'Superscript', exact: true }).click();
		await editor.getByRole('button', { name: 'Subscript', exact: true }).click();
		await expect(surface.locator('s')).toHaveText('Parity formatting');
		await expect(surface.locator('sub')).toHaveText('Parity formatting');
		await expect(surface.locator('sup')).toHaveCount(0);
		await expect(surface.locator('span[style*="background-color"]')).toHaveCSS(
			'background-color',
			'rgb(0, 255, 255)',
		);

		const formattedDownload = page.waitForEvent('download');
		await saveButton(page).click();
		const formatted = await formattedDownload;
		const formattedBytes = await readFile((await formatted.path())!);
		const formattedZip = await JSZip.loadAsync(formattedBytes);
		const formattedXml = await formattedZip.file('word/document.xml')!.async('string');
		expect(formattedXml).toContain('<w:strike');
		expect(formattedXml).toContain('<w:highlight w:val="cyan"');
		expect(formattedXml).toContain('<w:vertAlign w:val="subscript"');
		expect(formattedXml).not.toContain('<w:vertAlign w:val="superscript"');
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'format-roundtrip.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: formattedBytes,
		});
		await expect(fileNameLabel(page)).toHaveText('format-roundtrip.docx');
		await expect(surface.locator('s sub')).toHaveText('Parity formatting');
		await expect(surface.locator('sup')).toHaveCount(0);

		// Insert and structurally edit a 2x2 table, exercising history and read-only state.
		await newDocument(page);
		await surface.click();
		// Table tools are contextual: no Table tab until the caret is in a table.
		await expect(editor.getByRole('tab', { name: 'Table', exact: true })).toHaveCount(0);
		await insertTableOfSize(page);
		const table = surface.locator('table');
		await expect(table.locator('tr')).toHaveCount(2);
		await expect(table.locator('tr').first().locator('td')).toHaveCount(2);
		await table.locator('tr').nth(1).locator('td').nth(1).click();
		await editor.getByRole('tab', { name: 'Table', exact: true }).click();
		await editor.getByRole('button', { name: 'Insert row below', exact: true }).click();
		await expect(table.locator('tr')).toHaveCount(3);
		await editor.getByRole('tab', { name: 'Home', exact: true }).click();
		await editor.getByRole('button', { name: 'Undo', exact: true }).click();
		await expect(table.locator('tr')).toHaveCount(2);
		await editor.getByRole('button', { name: 'Redo', exact: true }).click();
		await expect(table.locator('tr')).toHaveCount(3);

		await table.locator('tr').nth(1).locator('td').nth(1).click();
		await editor.getByRole('tab', { name: 'Table', exact: true }).click();
		await editor.getByRole('button', { name: 'Insert column right', exact: true }).click();
		for (const row of await table.locator('tr').all())
			await expect(row.locator('td')).toHaveCount(3);
		await editor.getByRole('tab', { name: 'Home', exact: true }).click();
		await editor.getByRole('button', { name: 'Undo', exact: true }).click();
		for (const row of await table.locator('tr').all())
			await expect(row.locator('td')).toHaveCount(2);
		await editor.getByRole('button', { name: 'Redo', exact: true }).click();
		for (const row of await table.locator('tr').all())
			await expect(row.locator('td')).toHaveCount(3);

		await table.locator('tr').first().locator('td').nth(1).click();
		await editor.getByRole('tab', { name: 'Table', exact: true }).click();
		await editor.getByRole('button', { name: 'Delete column', exact: true }).click();
		await expect(table.locator('tr').first().locator('td')).toHaveCount(2);
		await table.locator('tr').nth(2).locator('td').first().click();
		await editor.getByRole('button', { name: 'Delete row', exact: true }).click();
		await expect(table.locator('tr')).toHaveCount(2);

		await setReadOnly(page, true);
		await expect(editor.getByRole('button', { name: 'Delete table', exact: true })).toBeDisabled();
		await setReadOnly(page, false);
		const tableDownload = page.waitForEvent('download');
		await saveButton(page).click();
		const tableFile = await tableDownload;
		const tableBytes = await readFile((await tableFile.path())!);
		const tableZip = await JSZip.loadAsync(tableBytes);
		const tableXml = await tableZip.file('word/document.xml')!.async('string');
		expect(tableXml).toContain('<w:tbl>');
		expect(tableXml.match(/<w:tr>/g)).toHaveLength(2);
		expect(tableXml.match(/<w:tc>/g)).toHaveLength(4);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'table-roundtrip.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: tableBytes,
		});
		await expect(fileNameLabel(page)).toHaveText('table-roundtrip.docx');
		await expect(surface.locator('table tr')).toHaveCount(2);
		for (const row of await surface.locator('table tr').all())
			await expect(row.locator('td')).toHaveCount(2);
		expect(errors).toEqual([]);
	});
}
