import { expect, test, type Page } from '@playwright/test';
import { newDocument, openSample, reveal, setReadOnly } from './helpers';

const editor = (page: Page) => page.locator('docx-editor');
const surface = (page: Page) => editor(page).locator('.ProseMirror');

/** Records file-command events and cancels them so no download or print dialog opens. */
async function recordFileCommands(page: Page) {
	await editor(page).evaluate((element) => {
		const seen: { command: string; cancelable: boolean }[] = [];
		(window as unknown as { fileCommands: typeof seen }).fileCommands = seen;
		element.addEventListener('file-command', (event) => {
			const custom = event as CustomEvent<{ command: string }>;
			seen.push({ command: custom.detail.command, cancelable: event.cancelable });
			event.preventDefault();
		});
	});
	return () =>
		page.evaluate(
			() =>
				(window as unknown as { fileCommands: { command: string; cancelable: boolean }[] })
					.fileCommands,
		);
}

async function startTyping(page: Page) {
	await openSample(page);
	await newDocument(page);
	await surface(page).click();
	await page.keyboard.type('Shortcut text');
}

test('Ctrl+F and Ctrl+H open find and replace, Escape returns to the document', async ({
	page,
}) => {
	await startTyping(page);
	await page.keyboard.press('Control+f');
	await expect(editor(page).getByLabel('Find text', { exact: true })).toBeFocused();
	await page.keyboard.press('Escape');
	await expect(editor(page).getByRole('region', { name: 'Find and replace' })).toBeHidden();
	await expect(surface(page)).toBeFocused();
	await page.keyboard.press('Control+h');
	await expect(editor(page).getByLabel('Replace with', { exact: true })).toBeFocused();
	await page.keyboard.press('Escape');
	await expect(surface(page)).toBeFocused();
});

test('Ctrl+F works with focus outside the document, inside the editor', async ({ page }) => {
	await startTyping(page);
	await editor(page).getByRole('tab', { name: 'Home', exact: true }).focus();
	await page.keyboard.press('Control+f');
	await expect(editor(page).getByLabel('Find text', { exact: true })).toBeFocused();
});

test('Ctrl+S and Ctrl+P go through cancelable file commands', async ({ page }) => {
	await startTyping(page);
	const commands = await recordFileCommands(page);
	await page.keyboard.press('Control+s');
	await page.keyboard.press('Control+p');
	expect(await commands()).toEqual([
		{ command: 'save', cancelable: true },
		{ command: 'print', cancelable: true },
	]);
});

test('F6 and Shift+F6 cycle ribbon, document and status bar; Alt and F10 focus the ribbon', async ({
	page,
}) => {
	await startTyping(page);
	const ribbonTab = editor(page).locator('[role="tab"][aria-selected="true"]');
	const statusButton = editor(page).locator('.dve-status button:not([hidden])').first();
	await page.keyboard.press('F6');
	await expect(statusButton).toBeFocused();
	await page.keyboard.press('F6');
	await expect(ribbonTab).toBeFocused();
	await page.keyboard.press('F6');
	await expect(surface(page)).toBeFocused();
	await page.keyboard.press('Shift+F6');
	await expect(ribbonTab).toBeFocused();
	await page.keyboard.press('Shift+F6');
	await expect(statusButton).toBeFocused();
	await page.keyboard.press('Escape');
	await expect(surface(page)).toBeFocused();
	await page.keyboard.press('F10');
	await expect(ribbonTab).toBeFocused();
	await page.keyboard.press('Escape');
	await page.keyboard.press('Alt');
	await expect(ribbonTab).toBeFocused();
});

test('Ctrl+/ and F1 open a localized shortcut help dialog that Escape closes', async ({ page }) => {
	await startTyping(page);
	const dialog = editor(page).getByRole('dialog', { name: 'Keyboard shortcuts' });
	await page.keyboard.press('Control+/');
	await expect(dialog).toBeVisible();
	await expect(dialog.getByRole('cell', { name: 'Find', exact: true })).toBeVisible();
	await expect(dialog.getByText('Ctrl+F', { exact: true })).toBeVisible();
	// ProseMirror's own formatting keys are listed because they exist.
	await expect(dialog.getByText('Ctrl+B', { exact: true })).toBeVisible();
	await expect(dialog.getByRole('button', { name: 'Close' })).toBeFocused();
	await page.keyboard.press('Escape');
	await expect(dialog).toBeHidden();
	await expect(surface(page)).toBeFocused();
	await editor(page).evaluate((element) => {
		(element as HTMLElement & { locale: string }).locale = 'fr';
	});
	await page.keyboard.press('F1');
	await expect(editor(page).getByRole('dialog', { name: 'Raccourcis clavier' })).toBeVisible();
	await page.keyboard.press('Escape');
});

test('Escape closes the comments panel and returns focus to the document', async ({ page }) => {
	await startTyping(page);
	await editor(page).getByRole('tab', { name: 'Review', exact: true }).click();
	const comments = editor(page)
		.locator('button[aria-label="Comments"]:not(.ribbon-overflow-button)')
		.first();
	await reveal(editor(page), comments);
	await comments.click();
	const panel = editor(page).getByRole('complementary', { name: 'Comments' });
	await expect(panel).toBeVisible();
	await editor(page).getByLabel('New comment', { exact: true }).focus();
	await page.keyboard.press('Escape');
	await expect(panel).toBeHidden();
	await expect(surface(page)).toBeFocused();
});

test('read-only disables the editing shortcut but keeps find and save', async ({ page }) => {
	await startTyping(page);
	await setReadOnly(page, true);
	const commands = await recordFileCommands(page);
	const search = editor(page).getByLabel('Find text', { exact: true });
	await surface(page).click();
	await page.keyboard.press('Control+h');
	await expect(search).toBeHidden();
	await page.keyboard.press('Control+f');
	await expect(search).toBeVisible();
	await page.keyboard.press('Control+s');
	expect(await commands()).toEqual([{ command: 'save', cancelable: true }]);
});
