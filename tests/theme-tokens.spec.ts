import { expect, test, type Page } from '@playwright/test';
import { openSample } from './helpers';

const bg = (page: Page, selector: string) =>
	page.locator(selector).evaluate((el) => getComputedStyle(el).backgroundColor);
const chromeBackground = (page: Page) => bg(page, 'docx-editor .dve-titlebar');
const paperBackground = (page: Page) => bg(page, 'docx-editor .dve-paper');
const setTheme = (page: Page, theme: string) =>
	page.locator('docx-editor').evaluate((el, value) => {
		(el as HTMLElement & { theme: string }).theme = value;
	}, theme);

test('theme property switches chrome tokens while the paper stays white', async ({ page }) => {
	await openSample(page);
	await setTheme(page, 'light');
	const lightChrome = await chromeBackground(page);
	const lightRibbon = await bg(page, 'docx-editor .dve-ribbon');
	await setTheme(page, 'dark');
	await expect(page.locator('docx-editor')).toHaveAttribute('theme', 'dark');
	expect(await chromeBackground(page)).not.toBe(lightChrome);
	expect(await bg(page, 'docx-editor .dve-ribbon')).not.toBe(lightRibbon);
	expect(await paperBackground(page)).toBe('rgb(255, 255, 255)');
	await setTheme(page, 'light');
	expect(await paperBackground(page)).toBe('rgb(255, 255, 255)');
	expect(await chromeBackground(page)).toBe(lightChrome);
});

test('auto follows the OS scheme and themeColors override the preset', async ({ page }) => {
	await page.emulateMedia({ colorScheme: 'dark' });
	await openSample(page);
	await setTheme(page, 'auto');
	const autoDark = await chromeBackground(page);
	await setTheme(page, 'dark');
	expect(await chromeBackground(page)).toBe(autoDark);
	await page.emulateMedia({ colorScheme: 'light' });
	await setTheme(page, 'auto');
	expect(await chromeBackground(page)).not.toBe(autoDark);
	await page.locator('docx-editor').evaluate((el) => {
		(el as HTMLElement & { themeColors: object }).themeColors = { background: 'rgb(1, 2, 3)' };
	});
	expect(await chromeBackground(page)).toBe('rgb(1, 2, 3)');
});

test('keyboard focus shows the ring on ribbon tabs', async ({ page }) => {
	await openSample(page);
	const tab = page.locator('docx-editor .ribbon-tabs [role="tab"]').first();
	await tab.focus();
	await page.keyboard.press('Tab');
	await page.keyboard.press('Shift+Tab');
	const outline = await tab.evaluate((el) => {
		const style = getComputedStyle(el);
		return { width: style.outlineWidth, style: style.outlineStyle };
	});
	expect(outline).toEqual({ width: '2px', style: 'solid' });
});
