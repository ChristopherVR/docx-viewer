import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { newDocument, reveal, saveButton } from './helpers';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

test('inherited headers edit together, unlink independently, relink and undo', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((el) => {
		const host = el as DocxEditorElement;
		const model = host.documentModel!;
		const geometry = {
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
				{ type: 'paragraph', id: 'p1', runs: [{ text: 'First section' }] },
				{ type: 'paragraph', id: 'p2', runs: [{ text: 'Second section' }] },
			],
			sections: [
				{
					...geometry,
					endsAtBlockId: 'p1',
					headers: {
						default: {
							partName: 'word/header1.xml',
							blocks: [{ type: 'paragraph', id: 'h', runs: [{ text: 'Company' }] }],
						},
					},
				},
				{ ...geometry, endsAtBlockId: 'p2' },
			],
		} as typeof model;
	});
	const body = editor.locator('.dve-paper > .ProseMirror');
	await body.locator('p').nth(1).click();
	await expect(editor.locator('.dve-header')).toContainText('Company');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const link = editor.getByRole('button', { name: 'Link to Previous', exact: true });
	await reveal(editor, link);
	await expect(link).toBeDisabled();
	await editor.locator('.dve-header [data-slot=default]').dblclick();
	await expect(link).toBeEnabled();
	await expect(link).toHaveAttribute('aria-pressed', 'true');
	await page.keyboard.press('End');
	await page.keyboard.type(' shared');
	await reveal(editor, link);
	await link.click();
	await expect(link).toHaveAttribute('aria-pressed', 'false');
	await editor.locator('.dve-header .ProseMirror').click();
	await page.keyboard.press('End');
	await page.keyboard.type(' separate');
	expect(
		await editor.evaluate(
			(el) => (el as DocxEditorElement).documentModel!.sections![0]!.headers!.default!.blocks[0],
		),
	).toMatchObject({ runs: [{ text: 'Company shared' }] });
	await reveal(editor, link);
	await link.click();
	await expect(link).toHaveAttribute('aria-pressed', 'true');
	await expect(editor.locator('.dve-header')).not.toContainText('separate');
	const linkedDownload = page.waitForEvent('download');
	await saveButton(page).click();
	await (await linkedDownload).saveAs('test-results/header-footer-linked.docx');
	await editor.locator('.dve-header .ProseMirror').press('Control+z');
	await expect(editor.locator('.dve-header')).toContainText('separate');
	await expect(link).toHaveAttribute('aria-pressed', 'false');
	await body.click();
	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	await (await pending).saveAs('test-results/header-footer-unlinked.docx');
	const zip = await JSZip.loadAsync(await readFile('test-results/header-footer-unlinked.docx'));
	expect(await zip.file('word/header1.xml')!.async('string')).not.toContain('separate');
	expect(await zip.file('word/header2.xml')!.async('string')).toContain('separate');
});
