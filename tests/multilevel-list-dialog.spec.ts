import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { computeListLabels, loadDocx } from '../packages/core/src/index';
import { restartFixture } from './support/restart-fixture';
import { fileInput, reveal, saveButton } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: define a multilevel list, undo, redo and export its restart and position settings`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		const bytes = await restartFixture(1);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'outline.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		const editor = page.locator('docx-editor');
		const body = editor.locator('.dve-paper > .ProseMirror');
		await body.click();
		await page.keyboard.press('Control+a');
		const menu = editor.getByRole('combobox', { name: 'Multilevel list', exact: true });
		await reveal(editor, menu);
		await menu.selectOption('define');
		const dialog = editor.getByRole('dialog', { name: 'Define New Multilevel List', exact: true });
		await expect(dialog).toBeVisible();
		await dialog.getByLabel('Start at', { exact: true }).fill('2');
		await dialog.getByLabel('Level to modify', { exact: true }).selectOption({ value: '2' });
		await dialog.getByLabel('Start at', { exact: true }).fill('4');
		await dialog.getByLabel('Enter formatting for number', { exact: true }).fill('%1.%3.');
		await dialog.getByLabel('Restart list after', { exact: true }).selectOption('0');
		await dialog.getByLabel('Number alignment', { exact: true }).selectOption('right');
		await dialog.getByLabel('Aligned at (inches)', { exact: true }).fill('0.25');
		await dialog.getByLabel('Text indent at (inches)', { exact: true }).fill('1');
		await dialog.getByLabel('Follow number with', { exact: true }).selectOption('space');
		if (framework === 'vanilla') {
			expect(
				await dialog
					.getByRole('button', { name: 'OK', exact: true })
					.evaluate((button) => button.getBoundingClientRect().bottom <= window.innerHeight),
			).toBe(true);
			await page.screenshot({ path: test.info().outputPath('multilevel-list-dialog.png') });
		}
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		const labels = () =>
			body
				.locator('p')
				.evaluateAll((nodes) => nodes.map((node) => node.getAttribute('data-list-label')?.trim()));
		const expected = ['2.', '1.', '2.4.', '2.5.', '3.', '3.6.', '2.', '3.7.'];
		await expect.poll(labels).toEqual(expected);
		await editor
			.locator('.dve-quick-access')
			.getByRole('button', { name: 'Undo', exact: true })
			.click();
		await expect.poll(labels).toEqual(['1.', '1.', '1.', '2.', '2.', '1.', '2.', '2.']);
		await editor
			.locator('.dve-quick-access')
			.getByRole('button', { name: 'Redo', exact: true })
			.click();
		await expect.poll(labels).toEqual(expected);
		const [download] = await Promise.all([page.waitForEvent('download'), saveButton(page).click()]);
		const path = test.info().outputPath('custom-list.docx');
		await download.saveAs(path);
		const output = new Uint8Array(await readFile(path));
		const loaded = await loadDocx(output);
		const computed = computeListLabels(loaded.model);
		expect(loaded.model.blocks.map((p) => computed.get(p.id)?.text)).toEqual(expected);
		const catalog = loaded.model.numberingCatalog!;
		expect(catalog.abstractNums['1']!.levels[2]).toMatchObject({
			start: 4,
			lvlRestart: 0,
			lvlText: '%1.%3.',
			lvlJc: 'right',
			indentLeftTwips: 1440,
			hangingTwips: 1080,
			suffix: 'space',
		});
		const original = await JSZip.loadAsync(bytes),
			zip = await JSZip.loadAsync(output);
		const originalXml = await original.file('word/numbering.xml')!.async('string');
		const outputXml = await zip.file('word/numbering.xml')!.async('string');
		expect(outputXml).toContain(originalXml.match(/<w:abstractNum\b[\s\S]*?<\/w:abstractNum>/)![0]);
	});
}
