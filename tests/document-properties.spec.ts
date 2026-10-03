import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from 'docx-web-component';
import { openSample } from './helpers';

test('File > Info edits the title and author, marks the document dirty and saves them', async ({
	page,
}) => {
	await openSample(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.dve-file-tab').click();
	await editor.locator('.dve-backstage-nav-item', { hasText: /^Info$/ }).click();
	const title = editor.getByRole('textbox', { name: 'Title', exact: true });
	await title.fill('Quarterly report');
	await title.press('Tab');
	await editor.getByRole('textbox', { name: 'Author', exact: true }).fill('Ann');
	await editor.getByRole('textbox', { name: 'Tags', exact: true }).fill('a, b');
	await editor.getByRole('textbox', { name: 'Tags', exact: true }).press('Tab');
	expect(await editor.evaluate((element) => (element as DocxEditorElement).dirty)).toBe(true);
	const bytes = await editor.evaluate(async (element) => {
		const blob = await (element as DocxEditorElement).save();
		return Array.from(new Uint8Array(await blob.arrayBuffer()));
	});
	const zip = await JSZip.loadAsync(Uint8Array.from(bytes));
	const xml = (await zip.file('docProps/core.xml')?.async('string')) ?? '';
	expect(xml).toContain('Quarterly report');
	expect(xml).toContain('Ann');
	expect(xml).toContain('a, b');
});
