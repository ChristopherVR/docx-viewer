import { expect, test } from '@playwright/test';
import { newDocument } from './helpers';

test('View > Multiple pages shows Print Layout pages side by side', async ({ page }) => {
	await page.setViewportSize({ width: 1700, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('First page');
	await page.keyboard.press('Control+Enter');
	await page.keyboard.type('Second page');
	await editor.locator('#dve-tab-view').click();
	await editor.getByRole('button', { name: 'Multiple pages' }).click();
	const pages = editor.locator('.dve-print-page');
	await expect(pages).toHaveCount(2);
	const [first, second] = await pages.evaluateAll((nodes) =>
		nodes.map((node) => {
			const box = node.getBoundingClientRect();
			return { top: Math.round(box.top), left: Math.round(box.left) };
		}),
	);
	expect(second!.top).toBe(first!.top);
	expect(second!.left).toBeGreaterThan(first!.left);
	// Any other zoom goes back to one column.
	await editor.getByRole('button', { name: 'Page width' }).click();
	const [a, b] = await pages.evaluateAll((nodes) =>
		nodes.map((node) => Math.round(node.getBoundingClientRect().top)),
	);
	expect(b).toBeGreaterThan(a!);
});
