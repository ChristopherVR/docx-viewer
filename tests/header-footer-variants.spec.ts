import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, reveal, saveButton } from './helpers';

test('first and even footer insertion, numbering, undo and Word export', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.dve-paper > .ProseMirror');
	await surface.click();
	await page.keyboard.type('Page one');
	await page.keyboard.press('Control+Enter');
	await page.keyboard.type('Page two');
	await page.keyboard.press('Control+Enter');
	await page.keyboard.type('Page three');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const footer = editor.getByRole('combobox', { name: 'Footer', exact: true });
	await reveal(editor, footer);
	await footer.selectOption('first');
	await expect(editor.locator('.dve-footer [data-slot=first]')).toBeVisible();
	await surface.click();
	await page.keyboard.press('Control+z');
	await expect(editor.locator('.dve-footer')).toHaveCount(0);
	expect(
		await editor.evaluate((el) =>
			Boolean((el as DocxEditorElement).documentModel!.sections![0]!.titlePage),
		),
	).toBe(false);
	await page.keyboard.press('Control+y');
	await editor.locator('.dve-footer [data-slot=first]').dblclick();
	await page.keyboard.type('First footer');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const numbers = editor.getByRole('combobox', { name: 'Page number', exact: true });
	await reveal(editor, numbers);
	await numbers.selectOption('bottom:center:pageOfTotal');
	await expect(editor.locator('.dve-footer [data-slot=first]')).toContainText('Page 1 of 1');
	await surface.click();
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await reveal(editor, footer);
	await footer.selectOption('even');
	await surface.click();
	await page.keyboard.press('Control+z');
	await expect(editor.locator('.dve-footer [data-slot=even]')).toHaveCount(0);
	expect(
		await editor.evaluate((el) =>
			Boolean((el as DocxEditorElement).documentModel!.evenAndOddHeaders),
		),
	).toBe(false);
	await page.keyboard.press('Control+y');
	await editor.locator('.dve-footer [data-slot=even]').dblclick();
	await page.keyboard.type('Even footer');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await reveal(editor, numbers);
	await numbers.selectOption('bottom:right');
	await expect(editor.locator('.dve-footer [data-slot=even]')).toContainText('Even footer');
	await expect(editor.locator('.dve-footer [data-slot=default]')).toHaveCount(0);
	await surface.click();
	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	await (await pending).saveAs('test-results/header-footer-variants.docx');
	const zip = await JSZip.loadAsync(await readFile('test-results/header-footer-variants.docx'));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toContain('<w:titlePg');
	expect(xml).toContain('w:type="first"');
	expect(xml).toContain('w:type="even"');
	expect(await zip.file('word/settings.xml')!.async('string')).toContain('<w:evenAndOddHeaders');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	const layout = editor.getByRole('combobox', { name: 'Layout view', exact: true });
	await reveal(editor, layout);
	await layout.selectOption('print');
	const pages = editor.locator('.dve-print-page');
	await expect(pages).toHaveCount(3);
	await expect(pages.first()).toContainText('First footer');
	await expect(pages.first()).toContainText('Page 1 of 3');
	await expect(pages.nth(1)).toContainText('Even footer');
	await expect(pages.nth(1)).not.toContainText('First footer');
	await expect(pages.nth(2)).not.toContainText('footer');
});
