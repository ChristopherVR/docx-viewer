import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import { fileNameLabel, newDocument } from './helpers';

const openFile = async (page: import('@playwright/test').Page) => {
	const editor = page.locator('docx-editor');
	await editor.locator('.dve-file-tab').click();
	return editor;
};
const navigate = (editor: import('@playwright/test').Locator, name: string) =>
	editor.locator('.dve-backstage-nav-item', { hasText: name }).click();

test.describe('File tab', () => {
	test.beforeEach(async ({ page }) => {
		await page.goto('/?framework=vanilla');
		await newDocument(page);
		await page.locator('docx-editor .ProseMirror').first().click();
		await page.keyboard.type('Hello file tab');
	});

	test('Info reports the document and its properties', async ({ page }) => {
		const editor = await openFile(page);
		await navigate(editor, 'Info');
		const content = editor.locator('.dve-backstage-content');
		await expect(content).toContainText('Document1.docx');
		await expect(content).toContainText('Unsaved changes');
		await expect(editor.locator('.dve-backstage-properties')).toContainText('3');
	});

	test('Home and Options work on a narrow screen and contain keyboard focus', async ({ page }) => {
		await page.setViewportSize({ width: 390, height: 844 });
		const editor = await openFile(page);
		await expect(editor.locator('.dve-backstage-content h2')).toHaveText('Home');
		await navigate(editor, 'Options');
		await editor.getByRole('checkbox', { name: 'Ruler', exact: true }).check();
		await expect(editor.locator('.dve-ruler')).toBeAttached();
		await editor.getByRole('checkbox', { name: 'Show paragraph marks', exact: true }).check();
		await expect(editor.locator('.dve-paper')).toHaveAttribute('data-show-marks', '');
		const last = editor.getByRole('checkbox', { name: 'Page thumbnails', exact: true });
		await last.focus();
		await page.keyboard.press('Tab');
		await expect(
			editor.getByRole('button', { name: 'Back to document', exact: true }),
		).toBeFocused();
		await page.keyboard.press('Shift+Tab');
		await expect(last).toBeFocused();
		const overflow = await editor
			.locator('.dve-backstage-content')
			.evaluate((el) => el.scrollWidth > el.clientWidth);
		expect(overflow).toBe(false);
		await page.keyboard.press('Escape');
		await expect(editor.locator('.dve-file-tab')).toBeFocused();
	});

	test('Save As renames the document and downloads it', async ({ page }) => {
		const editor = await openFile(page);
		await navigate(editor, 'Save As');
		await editor.getByLabel('File name', { exact: true }).fill('Quarterly report');
		const pending = page.waitForEvent('download');
		await editor.locator('.dve-backstage-content').getByRole('button', { name: 'Save' }).click();
		const download = await pending;
		expect(download.suggestedFilename()).toBe('Quarterly report.docx');
		await expect(fileNameLabel(page)).toHaveText('Quarterly report.docx');
		const zip = await JSZip.loadAsync(await readFile((await download.path())!));
		expect(await zip.file('word/document.xml')!.async('string')).toContain('Hello file tab');
	});

	test('Export writes plain text', async ({ page }) => {
		const editor = await openFile(page);
		await navigate(editor, 'Export');
		const pending = page.waitForEvent('download');
		await editor
			.locator('.dve-backstage-content')
			.getByRole('button', { name: 'Export as plain text' })
			.click();
		const download = await pending;
		expect(download.suggestedFilename()).toBe('Document1.txt');
		expect(await readFile((await download.path())!, 'utf8')).toBe('Hello file tab');
	});

	test('Options switches the display language and the review author', async ({ page }) => {
		const editor = await openFile(page);
		await navigate(editor, 'Options');
		await editor.getByLabel('Author name', { exact: true }).fill('Dana');
		await editor.getByLabel('Author name', { exact: true }).dispatchEvent('change');
		await expect(editor).toHaveAttribute('review-author', 'Dana');
		await editor.getByLabel('Display language', { exact: true }).selectOption('fr');
		await expect(editor).toHaveAttribute('locale', 'fr');
		await expect(editor.locator('.dve-backstage-content h2')).toHaveText('Options');
		await expect(editor.locator('.dve-backstage-nav-item', { hasText: 'Exporter' })).toBeVisible();
	});

	test('Escape closes the backstage and returns to the document', async ({ page }) => {
		const editor = await openFile(page);
		await expect(editor.locator('.dve-backstage')).toBeVisible();
		await page.keyboard.press('Escape');
		await expect(editor.locator('.dve-backstage')).toBeHidden();
	});
});
