import { expect, test } from '@playwright/test';
import { openSample } from './helpers';

test('File > Customize Ribbon hides and restores commands', async ({ page }) => {
	await page.setViewportSize({ width: 1700, height: 800 });
	await openSample(page);
	const editor = page.locator('docx-editor');
	const events = await editor.evaluate((element) => {
		(window as unknown as { seen: string[][] }).seen = [];
		element.addEventListener('ribbon-customize', (event) =>
			(window as unknown as { seen: string[][] }).seen.push([
				...(event as CustomEvent<string[]>).detail,
			]),
		);
		return true;
	});
	expect(events).toBe(true);
	await editor.locator('.dve-file-tab').click();
	await editor.locator('.dve-backstage-nav-item', { hasText: 'Customize Ribbon' }).click();
	await editor
		.locator('.dve-customize-row', { hasText: /^Bold$/ })
		.locator('input')
		.uncheck();
	await editor.getByRole('button', { name: 'Back to document' }).click();
	await expect(editor.getByRole('button', { name: 'Bold', exact: true })).toBeHidden();
	expect(
		await page.evaluate(() => (window as unknown as { seen: string[][] }).seen.at(-1)),
	).toEqual(['bold']);
	await editor.locator('.dve-file-tab').click();
	await editor.locator('.dve-backstage-nav-item', { hasText: 'Customize Ribbon' }).click();
	await editor.getByRole('button', { name: 'Reset all customizations' }).click();
	await editor.getByRole('button', { name: 'Back to document' }).click();
	await expect(editor.getByRole('button', { name: 'Bold', exact: true })).toBeVisible();
});
