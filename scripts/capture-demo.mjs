import { chromium } from '@playwright/test';
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({
	viewport: { width: 1440, height: 1150 },
	deviceScaleFactor: 1,
});
const errors = [];
page.on('pageerror', (error) => errors.push(error.message));
await page.goto('http://127.0.0.1:4181/');
await page.locator('docx-editor .ProseMirror').waitFor();
await page.screenshot({ path: 'artifacts/editor-desktop.png', fullPage: true });
await page.setViewportSize({ width: 390, height: 844 });
await page.screenshot({ path: 'artifacts/editor-mobile.png', fullPage: true });
console.log(
	JSON.stringify({
		errors,
		desktop: 'artifacts/editor-desktop.png',
		mobile: 'artifacts/editor-mobile.png',
	}),
);
await browser.close();
