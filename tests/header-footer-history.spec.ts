import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, reveal, saveButton } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: header insertion, typing and body edits share undo and redo`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.dve-paper > .ProseMirror');
		await surface.click();
		await page.keyboard.type('Body');
		await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
		const headerButton = editor.getByRole('combobox', { name: 'Header', exact: true });
		await reveal(editor, headerButton);
		await headerButton.selectOption('blank');
		const header = editor.locator('.dve-header');
		await expect(header).toBeVisible();
		await surface.click();
		await page.keyboard.press('Control+z');
		await expect(header).toHaveCount(0);
		await page.keyboard.press('Control+y');
		await expect(header).toBeVisible();
		await header.locator('[data-slot=default]').dblclick();
		const nested = header.locator('.ProseMirror');
		await expect(nested).toBeFocused();
		await page.keyboard.type('Company');
		await page.keyboard.press('Control+z');
		await expect(nested).toHaveText('');
		await page.keyboard.press('Control+y');
		await expect(nested).toHaveText('Company');
		await surface.click();
		await page.keyboard.press('End');
		await page.keyboard.type(' edited');
		await page.keyboard.press('Control+z');
		await expect(surface).toHaveText('Body');
		await expect(header).toContainText('Company');
		await page.keyboard.press('Control+z');
		await expect(header).not.toContainText('Company');
		await page.keyboard.press('Control+y');
		await expect(header).toContainText('Company');
	});
}

test('page numbers target the current section and its preview follows selection', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		const model = host.documentModel!;
		const section = {
			type: 'nextPage' as const,
			orientation: 'portrait' as const,
			pageWidthTwips: 12240,
			pageHeightTwips: 15840,
			marginTopTwips: 1440,
			marginLeftTwips: 1440,
			marginRightTwips: 1440,
			marginBottomTwips: 1440,
			columns: { count: 1, equalWidth: true },
		};
		host.documentModel = {
			...model,
			blocks: [
				{ type: 'paragraph', id: 'first', runs: [{ text: 'First section' }] },
				{ type: 'paragraph', id: 'second', runs: [{ text: 'Second section' }] },
			],
			sections: [
				{ ...section, endsAtBlockId: 'first' },
				{ ...section, endsAtBlockId: 'second' },
			],
		} as typeof model;
	});
	const paragraphs = editor.locator('.dve-paper > .ProseMirror > p');
	await paragraphs.nth(1).click({ position: { x: 12, y: 8 } });
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const menu = editor.getByRole('combobox', { name: 'Page number', exact: true });
	await reveal(editor, menu);
	await menu.selectOption('bottom:center');
	await expect
		.poll(() =>
			editor.evaluate((element) =>
				(element as DocxEditorElement).documentModel!.sections!.map(
					(section) => !!section.footers?.default,
				),
			),
		)
		.toEqual([false, true]);
	await expect(editor.locator('.dve-footer')).toBeVisible();
	await paragraphs.first().click({ position: { x: 12, y: 8 } });
	await expect(editor.locator('.dve-footer')).toHaveCount(0);
	await paragraphs.nth(1).click({ position: { x: 12, y: 8 } });
	await expect(editor.locator('.dve-footer')).toBeVisible();
	const promise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await promise;
	await download.saveAs('test-results/section-page-number.docx');
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml.match(/<w:footerReference/g)).toHaveLength(1);
	expect(xml.indexOf('<w:footerReference')).toBeGreaterThan(xml.indexOf('Second section'));
});
