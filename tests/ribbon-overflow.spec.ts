import { expect, test } from '@playwright/test';
import { openSample } from './helpers';

test.describe('ribbon overflow', () => {
	test.beforeEach(async ({ page }) => {
		await page.setViewportSize({ width: 1000, height: 700 });
		await openSample(page);
	});

	test('a narrow ribbon folds groups into dropdowns that still run their commands', async ({
		page,
	}) => {
		const editor = page.locator('docx-editor');
		const panel = editor.locator('.ribbon-panel[data-panel="Home"]');
		const fits = () => panel.evaluate((element) => element.scrollWidth <= element.clientWidth + 1);
		expect(await fits()).toBe(true);
		const folded = editor.locator('.ribbon-group[data-collapsed] > .ribbon-overflow-button');
		await expect(folded.first()).toBeVisible();
		await editor.locator('.ProseMirror p').first().click();
		await folded.filter({ hasText: 'Paragraph' }).click();
		await editor.getByRole('button', { name: 'Align center' }).click();
		await expect(editor.locator('.ProseMirror p').first()).toHaveCSS('text-align', 'center');
		// A command ends the dropdown, and the controls return to their group.
		await expect(editor.locator('.ribbon-overflow-panel')).toHaveCount(0);
	});

	test('widening the window unfolds the groups again', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.ribbon-group[data-collapsed]').first()).toBeAttached();
		await page.setViewportSize({ width: 1700, height: 700 });
		await expect(editor.locator('.ribbon-group[data-collapsed]')).toHaveCount(0);
		await expect(editor.getByRole('button', { name: 'Align center' })).toBeVisible();
	});

	test('Escape closes an open group dropdown', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await editor.locator('.ribbon-group[data-collapsed] > .ribbon-overflow-button').first().click();
		await expect(editor.locator('.ribbon-overflow-panel')).toHaveCount(1);
		await page.keyboard.press('Escape');
		await expect(editor.locator('.ribbon-overflow-panel')).toHaveCount(0);
	});
});
