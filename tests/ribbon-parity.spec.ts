import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { insertTableOfSize, newDocument, saveButton } from './helpers';

test.describe('Word-style ribbon', () => {
	test.beforeEach(async ({ page }) => {
		await page.goto('/?framework=vanilla');
		await newDocument(page);
	});

	test('Change Case, Select All and paragraph marks work from the Home tab', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await surface.click();
		await page.keyboard.type('hello world');
		await editor.getByRole('button', { name: 'Select all', exact: true }).click();
		await editor.getByRole('combobox', { name: 'Change case' }).selectOption('upper');
		await expect(surface).toContainText('HELLO WORLD');
		await editor.getByRole('combobox', { name: 'Change case' }).selectOption('upper');
		await expect(surface).toContainText('HELLO WORLD');
		await editor.getByRole('button', { name: 'Show paragraph marks', exact: true }).click();
		await expect(editor.locator('.dve-paper')).toHaveAttribute('data-show-marks', '');
		const mark = await surface
			.locator('p')
			.first()
			.evaluate((p) => getComputedStyle(p, '::after').content);
		expect(mark).toBe('"¶"');
	});

	test('Format Painter copies bold to the next selection', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		const painter = editor.getByRole('button', { name: 'Format painter', exact: true });
		await surface.click();
		await page.keyboard.type('alpha beta');
		// Keyboard selection is applied asynchronously; retry until the editor reports one (Cut enables).
		await expect(async () => {
			await page.keyboard.press('Control+Home');
			await page.keyboard.press('Control+Shift+ArrowRight');
			await expect(editor.getByRole('button', { name: 'Cut', exact: true })).toBeEnabled({
				timeout: 1000,
			});
		}).toPass({ timeout: 10000 });
		await editor.getByRole('button', { name: 'Bold', exact: true }).click();
		await expect(editor.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		await painter.click();
		await expect(painter).toHaveAttribute('aria-pressed', 'true');
		await page.keyboard.press('End');
		await page.keyboard.press('Control+Shift+ArrowLeft');
		await expect(painter).toHaveAttribute('aria-pressed', 'false');
		// 'alpha ' was bold; painting 'beta' extends the bold run over the whole text.
		await expect(surface.locator('strong')).toHaveText('alpha beta');
	});

	test('the Styles gallery lists styles and applies one to the paragraph', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await surface.click();
		await page.keyboard.type('A heading');
		const tile = editor.locator('.style-tile[data-style-id="Heading1"]');
		await expect(tile).toBeVisible();
		await tile.click();
		await expect(tile).toHaveAttribute('aria-pressed', 'true');
	});

	test('the ribbon collapses, peeks and shows KeyTips', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const panel = editor.locator('.ribbon-panel:not([hidden])');
		await expect(panel).toBeVisible();
		await editor.locator('.ribbon-collapse').click();
		await expect(panel).toBeHidden();
		await editor.locator('#dve-tab-insert').click();
		await expect(editor.locator('#dve-panel-insert')).toBeVisible();
		await page.keyboard.press('Escape');
		await expect(editor.locator('#dve-panel-insert')).toBeHidden();
		await editor.locator('.ribbon-collapse').click();
		await expect(editor.locator('#dve-panel-insert')).toBeVisible();

		await editor.locator('.ProseMirror').click();
		await page.keyboard.press('Alt');
		await expect(editor.locator('.dve-keytip').first()).toBeVisible();
		await page.keyboard.press('r');
		await expect(editor.locator('#dve-tab-review')).toHaveAttribute('aria-selected', 'true');
	});

	test('Insert, Review, View and Layout commands act on the document', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await surface.click();
		await page.keyboard.type('one two three');

		await editor.locator('#dve-tab-insert').click();
		await editor.getByRole('combobox', { name: 'Symbol' }).selectOption('©');
		await expect(surface).toContainText('one two three©');

		await editor.locator('#dve-tab-review').click();
		await editor.getByRole('button', { name: 'Word count', exact: true }).click();
		await expect(editor.locator('.dve-word-count')).toContainText('3');
		await page.keyboard.press('Escape');
		await expect(editor.locator('.dve-word-count')).toHaveCount(0);

		await editor.locator('#dve-tab-layout').click();
		const left = editor.getByLabel('Indent left', { exact: true });
		await left.fill('0.5');
		await left.dispatchEvent('change');
		await expect(surface.locator('p').first()).toHaveCSS('margin-left', '48px');

		await editor.locator('#dve-tab-view').click();
		await editor.getByRole('button', { name: 'Page width', exact: true }).click();
		const zoom = await editor
			.getByLabel('Document page', { exact: true })
			.evaluate((el) => Number(getComputedStyle(el).zoom));
		expect(zoom).toBeGreaterThan(0.1);
		await editor.getByRole('button', { name: 'Zoom to 100%', exact: true }).click();
		await expect(editor.getByLabel('Document page', { exact: true })).toHaveCSS('zoom', '1');
	});

	test('toggle buttons reflect the paragraph and the font boxes show the text', async ({
		page,
	}) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror').first();
		await surface.click();
		await page.keyboard.type('state test');
		const pressed = (name: string) =>
			editor.getByRole('button', { name, exact: true }).getAttribute('aria-pressed');
		expect(await pressed('Align left')).toBe('true');
		await editor.getByRole('button', { name: 'Align center', exact: true }).click();
		await expect(editor.getByRole('button', { name: 'Align center', exact: true })).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		expect(await pressed('Align left')).toBe('false');
		await editor.getByRole('button', { name: 'Bulleted list', exact: true }).click();
		await expect(
			editor.getByRole('button', { name: 'Bulleted list', exact: true }),
		).toHaveAttribute('aria-pressed', 'true');
		await expect(editor.getByRole('button', { name: 'Cut', exact: true })).toBeDisabled();
		await page.keyboard.press('Control+a');
		await expect(editor.getByRole('button', { name: 'Cut', exact: true })).toBeEnabled();
		await editor.getByLabel('Font family', { exact: true }).fill('Georgia');
		await page.keyboard.press('Enter');
		await expect(editor.getByLabel('Font family', { exact: true })).toHaveValue('Georgia');
		await editor.getByLabel('Font size', { exact: true }).fill('10.5');
		await page.keyboard.press('Enter');
		await expect(surface.locator('span').first()).toHaveCSS('font-size', '14px');
	});

	test('the Font and Paragraph dialogs apply their settings', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror').first();
		await surface.click();
		await page.keyboard.type('dialog text');
		await page.keyboard.press('Control+a');
		await expect(editor.getByRole('button', { name: 'Cut', exact: true })).toBeEnabled();
		await editor.getByRole('button', { name: 'Font settings' }).click();
		await editor.getByLabel('Small caps', { exact: true }).check();
		await editor.getByLabel('Underline style', { exact: true }).selectOption('double');
		await editor.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(editor.locator('.dve-font-dialog')).toBeHidden();
		await expect(surface.locator('[data-run-props]').first()).toHaveCSS(
			'font-variant-caps',
			'small-caps',
		);
		await expect(surface.locator('u')).toHaveCount(1);
		await editor.getByRole('button', { name: 'Paragraph settings' }).click();
		await editor.getByLabel('Left', { exact: true }).fill('1');
		await editor.getByLabel('Left', { exact: true }).dispatchEvent('input');
		await editor.getByLabel('Alignment', { exact: true }).selectOption('justify');
		await editor.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(surface.locator('p').first()).toHaveCSS('margin-left', '96px');
		await expect(surface.locator('p').first()).toHaveCSS('text-align', 'justify');
	});

	test('the Table tab is contextual', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror').first().click();
		await expect(editor.locator('#dve-tab-table')).toBeHidden();
		await insertTableOfSize(page);
		await editor.locator('.ProseMirror td').first().click();
		await expect(editor.locator('#dve-tab-table')).toBeVisible();
		await editor.locator('.ProseMirror').first().locator('p').first().click();
	});

	test('Insert > Page number adds a footer that is saved into the DOCX', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror').first().click();
		await page.keyboard.type('Body text');
		await editor.locator('#dve-tab-insert').click();
		await editor.getByRole('combobox', { name: 'Page number' }).selectOption('bottom:center');
		const footer = editor
			.locator('.dve-footer, [data-slot="footer"], .dve-header-footer-slot')
			.first();
		await expect(footer).toContainText('1');
		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const download = await pending;
		const zip = await JSZip.loadAsync(await readFile((await download.path())!));
		const footerXml = await zip.file('word/footer1.xml')!.async('string');
		expect(footerXml).toContain('PAGE');
		expect(footerXml).toContain('w:jc w:val="center"');
		expect(await zip.file('word/document.xml')!.async('string')).toContain('<w:footerReference');
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/word/footer1.xml');
	});

	test('a created footer survives body edits and undo', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror').first().click();
		await page.keyboard.type('Body text');
		await editor.locator('#dve-tab-insert').click();
		await editor.getByRole('combobox', { name: 'Page number' }).selectOption('bottom:right');
		await editor.locator('.ProseMirror').first().click();
		await page.keyboard.press('End');
		await page.keyboard.type(' more');
		await editor.getByRole('button', { name: 'Undo', exact: true }).click();
		const footerCount = () =>
			editor.evaluate((el) => {
				const model = (
					el as unknown as { documentModel: { sections?: Array<{ footers?: unknown }> } }
				).documentModel;
				return model.sections?.[0]?.footers ? 1 : 0;
			});
		expect(await footerCount()).toBe(1);
		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
		expect(await zip.file('word/footer1.xml')!.async('string')).toContain('PAGE');
	});

	test('Shading and Borders reach the saved DOCX', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror').first();
		await surface.click();
		await page.keyboard.type('Boxed and shaded');
		await editor.getByRole('button', { name: 'Shading options' }).click();
		await editor.getByRole('menuitem', { name: 'Cyan', exact: true }).click();
		await editor.getByRole('combobox', { name: 'Borders' }).selectOption('all');
		await expect(surface.locator('p').first()).toHaveCSS('background-color', 'rgb(0, 255, 255)');
		await expect(surface.locator('p').first()).toHaveCSS('border-bottom-style', 'solid');
		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('w:fill="00FFFF"');
		for (const side of ['top', 'left', 'bottom', 'right'])
			expect(xml).toContain(`<w:${side} w:val="single"`);
	});

	test('the expanded Styles gallery and the Styles pane apply styles', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror').first();
		await surface.click();
		await page.keyboard.type('Styled heading');
		await editor.getByRole('button', { name: 'More styles' }).click();
		const popover = editor.locator('.styles-popover');
		await expect(popover).toBeVisible();
		await expect(popover.locator('.style-tile[data-style-id="TOC1"]')).toHaveCount(0);
		await popover.locator('.style-tile[data-style-id="Heading2"]').click();
		await expect(popover).toHaveCount(0);
		await expect(editor.locator('.style-tile[data-style-id="Heading2"]').first()).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		await editor.getByRole('button', { name: 'Styles pane', exact: true }).click();
		const pane = editor.locator('.dve-styles-pane');
		await expect(pane).toBeVisible();
		await pane.locator('.style-tile[data-style-id="TOC1"]').click();
		await expect(pane.locator('.style-tile[data-style-id="TOC1"]')).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		await editor.getByRole('button', { name: 'Styles pane', exact: true }).click();
		await expect(pane).toBeHidden();
	});

	test('Find > Go to jumps to a heading', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror').first();
		await surface.click();
		await page.keyboard.type('Chapter one');
		await editor.locator('.style-tile[data-style-id="Heading1"]').first().click();
		await page.keyboard.press('Enter');
		await page.keyboard.type('Some later text');
		await editor.getByRole('combobox', { name: 'Find options' }).selectOption('goto');
		const panel = editor.locator('.go-to-panel');
		const items = panel.locator('.go-to-list [role="option"]');
		await expect(items).toHaveText(['Chapter one']);
		await items.click();
		await expect(panel).toHaveCount(0);
		// The caret is now at the start of the heading, so typing lands there.
		await page.keyboard.type('X');
		await expect(surface.locator('p').first()).toHaveText('XChapter one');
	});

	test('Insert > Bookmark and Cover page reach the saved DOCX', async ({ page }) => {
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror').first();
		await surface.click();
		await page.keyboard.type('Body paragraph');
		await editor.locator('#dve-tab-insert').click();
		await editor.getByRole('button', { name: 'Bookmark', exact: true }).click();
		await editor.getByLabel('Bookmark name', { exact: true }).fill('KeyPoint');
		await editor.locator('.dve-bookmark-dialog').getByRole('button', { name: 'Add' }).click();
		await expect(editor.locator('.dve-bookmark-dialog')).toBeHidden();
		await editor.getByRole('button', { name: 'Cover page', exact: true }).click();
		await expect(surface.locator('p').nth(1)).toHaveText('Document title');
		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('w:name="KeyPoint"');
		expect(xml).toContain('Document title');
		expect(xml).toContain('<w:pageBreakBefore/>');
	});

	test('the Page Setup dialog changes margins and paper and saves them', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror').first().click();
		await page.keyboard.type('Page setup text');
		await editor.locator('#dve-tab-layout').click();
		await editor.getByRole('button', { name: 'Page setup settings' }).click();
		const dialog = editor.locator('.dve-page-setup-dialog');
		await expect(dialog).toBeVisible();
		await dialog.getByLabel('Left', { exact: true }).fill('0.5');
		await dialog.getByLabel('Paper size', { exact: true }).selectOption('a4');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(dialog).toBeHidden();
		const pageWidth = await editor
			.getByLabel('Document page', { exact: true })
			.evaluate((el) => Math.round(parseFloat(getComputedStyle(el).width)));
		expect(pageWidth).toBe(794);
		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('w:w="11906"');
		expect(xml).toContain('w:left="720"');
	});

	test('View > Ruler shows the margins and follows the paragraph indent', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror').first().click();
		await page.keyboard.type('Indented paragraph');
		await editor.locator('#dve-tab-view').click();
		await editor.getByRole('button', { name: 'Ruler', exact: true }).click();
		const ruler = editor.locator('.dve-ruler');
		await expect(ruler).toBeVisible();
		const left = ruler.locator('.dve-ruler-marker-left');
		const before = await left.evaluate((el) => parseFloat((el as HTMLElement).style.left));
		await editor.locator('#dve-tab-layout').click();
		const spinner = editor.getByLabel('Indent left', { exact: true });
		await spinner.fill('1');
		await spinner.dispatchEvent('change');
		await expect
			.poll(() => left.evaluate((el) => parseFloat((el as HTMLElement).style.left)))
			.toBe(before + 96);
		await editor.locator('#dve-tab-view').click();
		await editor.getByRole('button', { name: 'Ruler', exact: true }).click();
		await expect(ruler).toHaveCount(0);
	});
});
