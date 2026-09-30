import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { loadDocx } from '../packages/core/src/index';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { newDocument, reveal, saveButton } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: Advanced Font controls apply, undo and export`, async ({ page }) => {
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			const host = element as DocxEditorElement;
			host.documentModel = {
				...host.documentModel!,
				blocks: [
					{
						type: 'paragraph',
						id: 'advanced',
						runs: [{ text: 'Advanced text', fontFamily: 'Courier New', fontSize: 12 }],
					},
				],
			};
		});
		const body = editor.locator('.dve-paper > .ProseMirror');
		await body.click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' changed');
		await page.keyboard.press('Control+a');
		const launcher = editor.getByRole('button', { name: 'Font settings', exact: true });
		await reveal(editor, launcher);
		await launcher.click();
		const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
		await dialog
			.getByRole('combobox', { name: 'Font style', exact: true })
			.selectOption('boldItalic');
		await dialog.getByRole('tab', { name: 'Advanced', exact: true }).click();
		await dialog.getByRole('spinbutton', { name: 'Scale (%)', exact: true }).fill('125');
		await dialog
			.getByRole('combobox', { name: 'Character spacing', exact: true })
			.selectOption('expanded');
		await dialog.getByRole('spinbutton', { name: 'By', exact: true }).fill('2');
		await dialog.getByRole('combobox', { name: 'Position', exact: true }).selectOption('lowered');
		await dialog.getByRole('spinbutton', { name: 'Position by (points)', exact: true }).fill('3');
		await dialog.getByRole('checkbox', { name: 'Kerning for fonts', exact: true }).check();
		await dialog.getByRole('spinbutton', { name: 'Points and above', exact: true }).fill('12');
		await dialog
			.getByRole('combobox', { name: 'Ligatures', exact: true })
			.selectOption('standardContextual');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(dialog).toBeHidden();
		const properties = () =>
			editor.evaluate((element) => {
				const p = (element as DocxEditorElement).documentModel!.blocks[0]!;
				if (p.type !== 'paragraph') throw new Error('Expected paragraph');
				return p.runs[0]!;
			});
		const expected = {
			bold: true,
			italic: true,
			textScalePercent: 125,
			characterSpacingTwips: 40,
			positionHalfPoints: -6,
			kerningHalfPoints: 24,
			ligatures: 'standardContextual',
		};
		expect(await properties()).toMatchObject(expected);
		await editor
			.locator('.dve-quick-access')
			.getByRole('button', { name: 'Undo', exact: true })
			.click();
		expect((await properties()).textScalePercent).toBeUndefined();
		expect((await properties()).bold).toBeUndefined();
		await expect(body).toHaveText('Advanced text changed');
		await editor
			.locator('.dve-quick-access')
			.getByRole('button', { name: 'Redo', exact: true })
			.click();
		expect(await properties()).toMatchObject(expected);
		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const download = await pending;
		const path = test.info().outputPath('advanced-font.docx');
		await download.saveAs(path);
		const { model } = await loadDocx(await readFile(path));
		const paragraph = model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected exported paragraph');
		expect(paragraph.runs[0]).toMatchObject(expected);
		if (framework === 'vanilla') {
			await launcher.click();
			await dialog.getByRole('tab', { name: 'Advanced', exact: true }).click();
			await dialog.screenshot({ path: test.info().outputPath('advanced-font.png') });
		}
	});
}

test('header Font changes use one global undo without undoing earlier header typing', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		const model = host.documentModel!;
		host.documentModel = {
			...model,
			sections: [
				{
					type: 'nextPage',
					orientation: 'portrait',
					pageWidthTwips: 12240,
					pageHeightTwips: 15840,
					marginTopTwips: 1440,
					marginBottomTwips: 1440,
					marginLeftTwips: 1440,
					marginRightTwips: 1440,
					columns: { count: 1, equalWidth: true },
					...model.sections?.[0],
					endsAtBlockId: model.blocks[0]!.id,
					headers: {
						default: {
							partName: 'word/header1.xml',
							blocks: [{ type: 'paragraph', id: 'header', runs: [{ text: 'Company' }] }],
						},
					},
				},
			],
		} as typeof model;
	});
	const header = editor.locator('.dve-header [data-slot=default]');
	await header.dblclick();
	await page.keyboard.press('Control+End');
	await page.keyboard.type(' changed');
	await page.keyboard.press('Control+a');
	await editor.getByRole('tab', { name: 'Home', exact: true }).click();
	const launcher = editor.getByRole('button', { name: 'Font settings', exact: true });
	await reveal(editor, launcher);
	await launcher.click();
	const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
	await dialog
		.getByRole('combobox', { name: 'Font style', exact: true })
		.selectOption('boldItalic');
	await dialog.getByRole('tab', { name: 'Advanced', exact: true }).click();
	await dialog.getByRole('spinbutton', { name: 'Scale (%)', exact: true }).fill('125');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const run = () =>
		editor.evaluate((element) => {
			const p = (element as DocxEditorElement).documentModel!.sections![0]!.headers!.default!
				.blocks[0]!;
			if (p.type !== 'paragraph') throw new Error('Expected header paragraph');
			return p.runs[0]!;
		});
	expect(await run()).toMatchObject({ bold: true, italic: true, textScalePercent: 125 });
	await editor
		.locator('.dve-quick-access')
		.getByRole('button', { name: 'Undo', exact: true })
		.click();
	expect((await run()).bold).toBeUndefined();
	expect((await run()).textScalePercent).toBeUndefined();
	await expect(header).toContainText('Company changed');
});
