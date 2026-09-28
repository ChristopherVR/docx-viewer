import { expect, test } from '@playwright/test';

const SHOT = process.env.LOCALE_SHOT;
const ready = (page: import('@playwright/test').Page) =>
	page.locator('html[data-demo-ready="true"]').waitFor({ state: 'attached' });

test.describe('demo interface language', () => {
	test('?locale= maps region variants and localizes the editor chrome', async ({ page }) => {
		await page.goto('/?locale=de-DE');
		await ready(page);
		await expect(page.locator('#locale-select')).toHaveValue('de');
		await page.locator('#sample').click();
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.ProseMirror')).toContainText('Document title');
		await expect(editor).toHaveAttribute('locale', 'de');
		await expect(editor).toHaveAttribute('aria-label', 'Dokumenteditor');
		await expect(editor.getByRole('button', { name: 'Fett', exact: true })).toBeVisible();
		await expect(editor.getByRole('tab', { name: 'Einfügen' })).toBeVisible();
		if (SHOT) await page.screenshot({ path: SHOT });
	});

	const picked: Array<[string, string, string]> = [
		['es', 'Negrita', 'Insertar'],
		['zh-CN', '加粗', '插入'],
		['fr', 'Gras', 'Insérer'],
	];
	for (const [locale, bold, insert] of picked)
		test(`the landing-page picker starts the editor in ${locale}`, async ({ page }) => {
			await page.goto('/');
			await ready(page);
			await page.locator('#locale-select').selectOption(locale);
			await expect(page).toHaveURL(new RegExp(`locale=${locale}`));
			await page.locator('#sample').click();
			const editor = page.locator('docx-editor');
			await expect(editor).toHaveAttribute('locale', locale);
			await expect(editor.getByRole('button', { name: bold, exact: true })).toBeVisible();
			await expect(editor.getByRole('tab', { name: insert })).toBeVisible();
		});
});
