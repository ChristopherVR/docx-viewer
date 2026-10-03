import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { fileInput, saveButton } from './helpers';

async function fixture() {
	const zip = new JSZip();
	const cell = (text: string, properties = '') =>
		`<w:tc><w:tcPr>${properties}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:test"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellMar><w:top w:w="72" w:type="dxa"/><w:left w:w="216" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell('A', '<w:tcW w:w="3000" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="4"/></w:tcBorders><w:shd w:fill="ABCDEF"/><w:tcMar><w:top w:w="144" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tcMar><ext:keep ext:value="yes"/>')}${cell('B', '<w:tcMar><w:left w:w="0" w:type="dxa"/><w:right w:w="432" w:type="dxa"/></w:tcMar>')}</w:tr><w:tr>${cell('C')}${cell('D')}</w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla']) {
	test(`${framework}: cell margin overrides edit, undo and reopen without altering other cells`, async ({
		page,
	}) => {
		page.on('dialog', (dialog) => void dialog.accept());
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'cell-margins.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await fixture(),
		});
		const editor = page.locator('docx-editor');
		const cells = editor.locator('.ProseMirror td');
		await expect(cells.first()).toHaveCSS('padding-top', '9.6px');
		await expect(cells.first()).toHaveCSS('padding-left', '14.4px');
		await expect(cells.nth(1)).toHaveCSS('padding-left', '0px');
		const open = async () => {
			await cells.first().click();
			await editor.getByRole('tab', { name: 'Table', exact: true }).click();
			await editor.getByRole('button', { name: 'Table properties', exact: true }).click();
			return editor.getByRole('dialog', { name: 'Table properties', exact: true });
		};
		let dialog = await open();
		await expect(dialog.getByRole('spinbutton', { name: 'Cell top', exact: true })).toHaveValue(
			'0.1',
		);
		await dialog.getByRole('spinbutton', { name: 'Cell left', exact: true }).fill('0.25');
		await dialog.getByRole('spinbutton', { name: 'Cell top', exact: true }).fill('0');
		await dialog.getByRole('spinbutton', { name: 'Top', exact: true }).fill('0.1');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(cells.first()).toHaveCSS('padding-left', '24px');
		await expect(cells.first()).toHaveCSS('padding-top', '0px');
		await expect(cells.nth(1)).toHaveCSS('padding-top', '9.6px');
		await expect(cells.nth(1)).toHaveCSS('padding-left', '0px');
		await page.keyboard.press('Control+z');
		await expect(cells.first()).toHaveCSS('padding-top', '9.6px');
		await expect(cells.first()).toHaveCSS('padding-left', '14.4px');
		await expect(cells.nth(1)).toHaveCSS('padding-top', '4.8px');
		await page.keyboard.press('Control+y');
		await expect(cells.first()).toHaveCSS('padding-left', '24px');
		const save = async () => {
			const pending = page.waitForEvent('download');
			await saveButton(page).click();
			return readFile((await (await pending).path())!);
		};
		let bytes = await save();
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<ext:keep ext:value="yes"/>');
		expect(xml).toContain('<w:tcW w:w="3000" w:type="dxa"/>');
		expect(xml).toContain('w:fill="ABCDEF"');
		expect(xml).toContain('<w:bottom w:val="single" w:sz="4"/>');
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'edited.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: bytes,
		});
		await expect(cells.first()).toHaveCSS('padding-left', '24px');
		dialog = await open();
		await dialog.getByRole('checkbox', { name: 'Use table defaults', exact: true }).check();
		await expect(dialog.getByRole('spinbutton', { name: 'Cell top', exact: true })).toBeDisabled();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(cells.first()).toHaveCSS('padding-top', '9.6px');
		await expect(cells.first()).toHaveCSS('padding-left', '14.4px');
		await expect(cells.nth(1)).toHaveCSS('padding-left', '0px');
		bytes = await save();
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'inherited.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: bytes,
		});
		await expect(cells.first()).toHaveCSS('padding-top', '9.6px');
		await expect(cells.first()).toHaveCSS('padding-left', '14.4px');
		await expect(cells.nth(1)).toHaveCSS('padding-right', '28.8px');
	});
}
