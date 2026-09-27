import { expect, test } from '@playwright/test';
import { openSample, saveStateLabel } from './helpers';

test('appearance changes from another window preserve the open document', async ({
	page,
	context,
}) => {
	await openSample(page);
	const surface = page.locator('docx-editor .ProseMirror');
	await expect(surface).toContainText('Document title');
	await surface.locator('p').last().click();
	await page.keyboard.press('End');
	await page.keyboard.type(' Theme synchronization keeps this edit.');
	const controls = await context.newPage();
	await controls.goto('/');
	const themeBefore = await controls.locator('html').getAttribute('data-theme');
	await controls.locator('#theme-toggle').click();
	await expect(page.locator('html')).toHaveAttribute(
		'data-theme',
		themeBefore === 'dark' ? 'light' : 'dark',
	);
	await expect(surface).toContainText('Theme synchronization keeps this edit.');
	await expect(saveStateLabel(page)).toHaveText('Unsaved changes');
	await controls.close();
});
