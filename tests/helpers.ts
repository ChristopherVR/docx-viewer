import { expect, type Locator, type Page } from '@playwright/test';

/** Browser-contract helpers for the demo host and the editor's built-in Word chrome. */
const editor = (page: Page) => page.locator('docx-editor');
const landingVisible = (page: Page) => page.locator('#landing').isVisible();
/** Waits until the demo script has wired the landing page's buttons. */
const demoReady = (page: Page) =>
	page.locator('html[data-demo-ready="true"]').waitFor({ state: 'attached' });

/** Opens the demo in `framework` and loads the sample document from the landing page. */
export async function openSample(page: Page, framework = 'vanilla') {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto(`/?framework=${framework}`);
	await demoReady(page);
	await page.locator('#sample').click();
	await expect(editor(page).locator('.ProseMirror')).toContainText('Document title');
}

/** Starts a blank document from the landing page, or through File > New in the editor. */
export async function newDocument(page: Page) {
	page.on('dialog', (dialog) => void dialog.accept());
	if (await landingVisible(page)) {
		await demoReady(page);
		await page.locator('#blank').click();
	} else {
		await editor(page).locator('.dve-file-tab').click();
		await editor(page).locator('.dve-backstage-nav-item', { hasText: 'New' }).click();
		await editor(page).getByRole('button', { name: 'Blank document' }).click();
	}
	await expect(editor(page).locator('.ProseMirror')).toBeVisible();
}

/** The file input that opens documents: the landing page's before an editor exists, else the editor's. */
export async function fileInput(page: Page) {
	await demoReady(page);
	return (await landingVisible(page))
		? page.locator('#landing-file')
		: editor(page).locator('input.dve-file-input');
}

export const saveButton = (page: Page) =>
	editor(page).locator('.dve-quick-access').getByRole('button', { name: 'Save', exact: true });

/** File > Export > Save a copy as DOCX. */
export async function saveCopyAsDocx(page: Page) {
	await editor(page).locator('.dve-file-tab').click();
	await editor(page).locator('.dve-backstage-nav-item', { hasText: 'Export' }).click();
	await editor(page)
		.locator('.dve-backstage-content')
		.getByRole('button', { name: 'Save a copy as DOCX' })
		.click();
}

export const fileNameLabel = (page: Page) => editor(page).locator('.dve-filename');
export const saveStateLabel = (page: Page) => editor(page).locator('.dve-save-state');

/** Switches the title bar's Editing/Viewing mode. */
export async function setReadOnly(page: Page, readOnly: boolean) {
	// Located by class: its accessible name follows the editor's display locale.
	await editor(page)
		.locator('.dve-mode-select')
		.selectOption(readOnly ? 'viewing' : 'editing');
}

/** Insert > Table: opens the size picker and picks `columns` x `rows` from its grid. */
export async function insertTableOfSize(page: Page, rows = 2, columns = 2) {
	const editor = page.locator('docx-editor');
	await editor.locator('#dve-tab-insert').click();
	await editor.getByRole('button', { name: 'Insert table', exact: true }).click();
	await editor
		.locator(`.table-picker [role="gridcell"][data-rows="${rows}"][data-columns="${columns}"]`)
		.click();
}

/**
 * A narrow ribbon folds groups into dropdown buttons, as Word does. Opens whichever of them holds
 * `target` (a no-op when it is already on the ribbon) so the test can use the control.
 */
export async function reveal(editor: Locator, target: Locator) {
	if (await target.first().isVisible()) return;
	const buttons = editor.locator('.ribbon-panel:not([hidden]) .ribbon-overflow-button');
	for (let index = 0; index < (await buttons.count()); index++) {
		await buttons.nth(index).click();
		if (await target.first().isVisible()) return;
		await editor.page().keyboard.press('Escape');
	}
}
