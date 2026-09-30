import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { newDocument, reveal, setReadOnly, saveButton } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: header/footer context navigates empty stories without inserting until typing`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((el) => {
			const host = el as DocxEditorElement;
			const model = host.documentModel!;
			const section = {
				type: 'nextPage' as const,
				orientation: 'portrait' as const,
				pageWidthTwips: 12240,
				pageHeightTwips: 15840,
				marginTopTwips: 1440,
				marginLeftTwips: 1440,
				marginRightTwips: 1440,
				marginBottomTwips: 1440,
				columns: { count: 1, equalWidth: true },
			};
			host.documentModel = {
				...model,
				blocks: [
					{ type: 'paragraph', id: 'a', runs: [{ text: 'First', bold: true }] },
					{ type: 'paragraph', id: 'b', runs: [{ text: 'Second' }] },
				],
				sections: [
					{
						...section,
						endsAtBlockId: 'a',
						headers: {
							default: {
								partName: 'word/header1.xml',
								blocks: [{ type: 'paragraph', id: 'h', runs: [{ text: 'Company' }] }],
							},
						},
					},
					{ ...section, endsAtBlockId: 'b' },
				],
			} as typeof model;
		});
		const tab = editor.getByRole('tab', { name: 'Header & Footer', exact: true });
		await expect(tab).toBeHidden();
		await editor.locator('.dve-header [data-slot=default]').dblclick();
		await expect(tab).toHaveAttribute('aria-selected', 'true');
		await editor.getByRole('tab', { name: 'Home', exact: true }).click();
		await expect(editor.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute(
			'aria-pressed',
			'false',
		);
		await tab.click();
		const panel = editor.getByRole('tabpanel', { name: 'Header & Footer', exact: true });
		await expect(
			panel.getByRole('button', { name: 'Previous Section', exact: true }),
		).toBeDisabled();
		const next = panel.getByRole('button', { name: 'Next Section', exact: true });
		await reveal(editor, next);
		await next.click();
		await expect(next).toBeDisabled();
		const headerDistance = panel.getByRole('spinbutton', {
			name: 'Header from Top (inches)',
			exact: true,
		});
		const footerDistance = panel.getByRole('spinbutton', {
			name: 'Footer from Bottom (inches)',
			exact: true,
		});
		await expect(headerDistance).toHaveValue('0.5');
		await headerDistance.fill('0.75');
		await headerDistance.press('Tab');
		await footerDistance.fill('0.875');
		await footerDistance.press('Tab');
		expect(
			await editor.evaluate((el) =>
				(el as DocxEditorElement).documentModel!.sections!.map((s) => [
					s.headerDistanceTwips ?? 720,
					s.footerDistanceTwips ?? 720,
				]),
			),
		).toEqual([
			[720, 720],
			[1080, 1260],
		]);
		await editor.locator('.dve-header .ProseMirror').press('Control+z');
		await expect(footerDistance).toHaveValue('0.5');
		await editor.locator('.dve-header .ProseMirror').press('Control+y');
		await expect(footerDistance).toHaveValue('0.875');
		if (framework === 'vanilla')
			await page.screenshot({ path: 'test-results/header-footer-toolbar.png' });
		await expect(
			panel.getByRole('button', { name: 'Link to Previous', exact: true }),
		).toHaveAttribute('aria-pressed', 'true');
		await panel.getByRole('button', { name: 'Go to Footer', exact: true }).click();
		await expect(editor.locator('.dve-footer .ProseMirror')).toBeFocused();
		expect(
			await editor.evaluate((el) =>
				(el as DocxEditorElement).documentModel!.sections!.some((s) => !!s.footers?.default),
			),
		).toBe(false);
		await panel.getByRole('button', { name: 'Close Header and Footer', exact: true }).click();
		await expect(tab).toBeHidden();
		await expect(editor.locator('.dve-footer')).toHaveCount(0);
		await expect(editor.locator('.dve-paper > .ProseMirror')).toBeFocused();
		await editor.locator('.dve-header [data-slot=default]').dblclick();
		await panel.getByRole('button', { name: 'Go to Footer', exact: true }).click();
		await page.keyboard.type('Section two footer');
		if (framework === 'vanilla') {
			const pending = page.waitForEvent('download');
			await saveButton(page).click();
			await (await pending).saveAs('test-results/header-footer-navigation.docx');
		}
		expect(
			await editor.evaluate((el) =>
				(el as DocxEditorElement).documentModel!.sections!.map((s) => !!s.footers?.default),
			),
		).toEqual([false, true]);
		await page.keyboard.press('Control+z');
		await expect(editor.locator('.dve-footer .ProseMirror')).toHaveText('');
		expect(
			await editor.evaluate((el) =>
				(el as DocxEditorElement).documentModel!.sections!.some((s) => !!s.footers?.default),
			),
		).toBe(false);
		await setReadOnly(page, true);
		await expect(tab).toBeHidden();
		await expect(editor.locator('.dve-footer')).toHaveCount(0);
		await setReadOnly(page, false);
		await editor.locator('.dve-paper > .ProseMirror').press('Control+y');
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		const layout = editor.getByRole('combobox', { name: 'Layout view', exact: true });
		await reveal(editor, layout);
		await layout.selectOption('print');
		const pages = editor.locator('.dve-print-page');
		await expect(pages).toHaveCount(2);
		await expect(pages.first().locator('.dve-print-header')).toHaveCSS('top', '48px');
		await expect(pages.nth(1).locator('.dve-print-header')).toHaveCSS('top', '72px');
		await expect(pages.nth(1).locator('.dve-print-footer')).toHaveCSS('bottom', '84px');
	});
}
