import { expect, test, type Locator } from '@playwright/test';
import { openSample } from './helpers';

/**
 * Taps Alt as a script does. A real Alt tap also moves Chromium's own focus to its menu bar on
 * Windows, which is not what the editor's shortcut handling is under test for.
 */
const tapAlt = (editor: Locator) =>
	editor.evaluate((element) => {
		const target = element.shadowRoot!.querySelector('.ProseMirror')!;
		for (const type of ['keydown', 'keyup'])
			target.dispatchEvent(new KeyboardEvent(type, { key: 'Alt', bubbles: true, composed: true }));
	});

test.describe('command KeyTips', () => {
	test.beforeEach(async ({ page }) => {
		await page.setViewportSize({ width: 1700, height: 800 });
		await openSample(page);
	});

	test('Alt, H, 1 makes the selection bold, as in Word', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const paragraph = editor.locator('.ProseMirror p').nth(2);
		await page.evaluate(() => document.fonts.ready);
		// A late layout pass can replace the selection under a slow machine; retry the whole chord.
		await expect(async () => {
			await page.keyboard.press('Escape');
			await paragraph.click({ clickCount: 3 });
			await tapAlt(editor);
			await page.keyboard.press('h');
			const tips = editor.locator('.dve-keytip-command');
			await expect(tips.first()).toBeVisible({ timeout: 2000 });
			expect(await tips.allTextContents()).toContain('1');
			await page.keyboard.press('1');
			await expect(editor.locator('.dve-keytip-command')).toHaveCount(0);
			await expect(paragraph.locator('strong, b, [style*="font-weight"]').first()).toBeAttached({
				timeout: 2000,
			});
		}).toPass({ timeout: 15000 });
	});

	test('no command tip is the start of another, and Escape leaves the tips', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror p').first().click();
		await tapAlt(editor);
		await page.keyboard.press('n');
		const tips = await editor.locator('.dve-keytip-command').allTextContents();
		expect(tips.length).toBeGreaterThan(5);
		for (const a of tips)
			for (const b of tips) if (a !== b) expect(b.startsWith(a), `${a} / ${b}`).toBe(false);
		await page.keyboard.press('Escape');
		await expect(editor.locator('.dve-keytip-command')).toHaveCount(0);
	});
});
