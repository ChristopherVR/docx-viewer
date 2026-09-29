import { expect, test, type Page } from '@playwright/test';
import { newDocument, openSample, setReadOnly, insertTableOfSize } from './helpers';

const editor = (page: Page) => page.locator('docx-editor');
const surface = (page: Page) => editor(page).locator('.ProseMirror');
const menu = (page: Page) => editor(page).getByRole('menu');
const item = (page: Page, name: string) => menu(page).getByRole('menuitem', { name, exact: true });

async function startTyping(page: Page, text = 'Context text') {
	await openSample(page);
	await newDocument(page);
	await surface(page).click();
	await page.keyboard.type(text);
}

async function insertTable(page: Page) {
	await insertTableOfSize(page);
	await expect(surface(page).locator('table')).toHaveCount(1);
}

test('right click opens a menu with clipboard, link and comment items but no table items', async ({
	page,
}) => {
	await startTyping(page);
	await surface(page).click({ button: 'right' });
	await expect(menu(page)).toBeVisible();
	await expect(menu(page).getByRole('menuitem')).toHaveText([
		'Cut',
		'Copy',
		'Paste',
		'Insert link',
		'Add comment',
	]);
	await expect(item(page, 'Cut')).toHaveAttribute('aria-disabled', 'true');
	await expect(item(page, 'Insert link')).not.toHaveAttribute('aria-disabled', 'true');
	await expect(item(page, 'Insert row above')).toHaveCount(0);
});

test('Escape, outside click and scroll close the menu; Escape restores document focus', async ({
	page,
}) => {
	await startTyping(page);
	await surface(page).click({ button: 'right' });
	await expect(menu(page)).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(menu(page)).toBeHidden();
	await expect(surface(page)).toBeFocused();
	await surface(page).click({ button: 'right' });
	await expect(menu(page)).toBeVisible();
	await editor(page).locator('.dve-status').click();
	await expect(menu(page)).toBeHidden();
	await surface(page).click({ button: 'right' });
	await expect(menu(page)).toBeVisible();
	await editor(page)
		.locator('.dve-canvas')
		.evaluate((canvas) => {
			canvas.scrollTop = 1;
			canvas.dispatchEvent(new Event('scroll'));
		});
	await expect(menu(page)).toBeHidden();
});

test('Shift+F10 and the ContextMenu key open it; arrows, Home and End navigate', async ({
	page,
}) => {
	await startTyping(page);
	await page.keyboard.press('Shift+F10');
	await expect(menu(page)).toBeVisible();
	// Cut and Copy are disabled without a selection, so Paste is the first enabled item.
	await expect(item(page, 'Paste')).toBeFocused();
	await page.keyboard.press('ArrowDown');
	await expect(item(page, 'Insert link')).toBeFocused();
	await page.keyboard.press('ArrowDown');
	await expect(item(page, 'Add comment')).toBeFocused();
	await page.keyboard.press('ArrowDown');
	await expect(item(page, 'Paste')).toBeFocused();
	await page.keyboard.press('ArrowUp');
	await expect(item(page, 'Add comment')).toBeFocused();
	await page.keyboard.press('Home');
	await expect(item(page, 'Paste')).toBeFocused();
	await page.keyboard.press('End');
	await expect(item(page, 'Add comment')).toBeFocused();
	await page.keyboard.press('Escape');
	await expect(menu(page)).toBeHidden();
	await page.keyboard.press('ContextMenu');
	await expect(menu(page)).toBeVisible();
	await page.keyboard.press('Home');
	await page.keyboard.press('ArrowDown');
	await page.keyboard.press('Enter');
	await expect(menu(page)).toBeHidden();
	await expect(editor(page).getByRole('dialog', { name: /link/i })).toBeVisible();
});

test('table actions appear only in a table and edit it', async ({ page }) => {
	await openSample(page);
	await newDocument(page);
	await surface(page).click();
	await insertTable(page);
	const rows = surface(page).locator('tr');
	const before = await rows.count();
	await surface(page).locator('td').first().click({ button: 'right' });
	await expect(menu(page)).toBeVisible();
	for (const name of [
		'Insert row above',
		'Insert row below',
		'Insert column left',
		'Insert column right',
		'Delete row',
		'Delete column',
		'Delete table',
	])
		await expect(item(page, name)).toBeVisible();
	await item(page, 'Insert row below').click();
	await expect(menu(page)).toBeHidden();
	await expect(rows).toHaveCount(before + 1);
	await surface(page).locator('td').first().click({ button: 'right' });
	await item(page, 'Delete table').click();
	await expect(surface(page).locator('table')).toHaveCount(0);
});

test('the menu stays inside the viewport when opened at the bottom-right corner', async ({
	page,
}) => {
	await startTyping(page);
	const viewport = page.viewportSize()!;
	await surface(page).evaluate((element) => {
		element.dispatchEvent(
			new MouseEvent('contextmenu', {
				bubbles: true,
				cancelable: true,
				clientX: window.innerWidth - 2,
				clientY: window.innerHeight - 2,
			}),
		);
	});
	await expect(menu(page)).toBeVisible();
	const box = (await menu(page).boundingBox())!;
	expect(box.x).toBeGreaterThanOrEqual(0);
	expect(box.y).toBeGreaterThanOrEqual(0);
	expect(box.x + box.width).toBeLessThanOrEqual(viewport.width);
	expect(box.y + box.height).toBeLessThanOrEqual(viewport.height);
});

test('read-only leaves only Copy enabled', async ({ page }) => {
	await startTyping(page);
	await page.keyboard.press('Control+a');
	await setReadOnly(page, true);
	await surface(page).click({ button: 'right' });
	await expect(menu(page)).toBeVisible();
	const enabled = menu(page).locator('[role="menuitem"]:not([aria-disabled])');
	await expect(enabled).toHaveText(['Copy']);
});

test.describe('clipboard', () => {
	test.use({ permissions: ['clipboard-read', 'clipboard-write'] });
	test('copy and paste through the menu', async ({ page }) => {
		await startTyping(page, 'abc');
		await page.keyboard.press('Control+a');
		await surface(page).click({ button: 'right' });
		await item(page, 'Copy').click();
		await page.keyboard.press('End');
		await surface(page).click({ button: 'right' });
		await item(page, 'Paste').click();
		await expect(surface(page)).toContainText('abcabc');
	});
});

test('labels are localized', async ({ page }) => {
	await startTyping(page);
	await editor(page).evaluate((element) => {
		(element as HTMLElement & { locale: string }).locale = 'fr';
	});
	await surface(page).click({ button: 'right' });
	await expect(menu(page)).toHaveAccessibleName('Menu contextuel');
	await expect(menu(page).getByRole('menuitem', { name: 'Coller', exact: true })).toBeVisible();
	await expect(menu(page).getByRole('menuitem', { name: 'Couper', exact: true })).toBeVisible();
});
