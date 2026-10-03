import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { fileInput, saveButton } from './helpers';

async function fixture() {
	const zip = new JSZip();
	const cell = (text: string, properties: string) =>
		`<w:tc><w:tcPr>${properties}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ext="urn:test"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="1440" w:hRule="atLeast"/></w:trPr>${cell('A', '<w:tcW w:w="3000" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="4"/></w:tcBorders><w:shd w:fill="ABCDEF"/><w:tcMar><w:left w:w="144" w:type="dxa"/></w:tcMar><w:vAlign w:val="bottom" ext:keep="yes"/><ext:keep ext:value="yes"/>')}${cell('B', '<w:vAlign w:val="center"/>')}</w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla']) {
	test(`${framework}: selected cell alignment edits, resets and reopens with source metadata`, async ({
		page,
	}) => {
		page.on('dialog', (dialog) => void dialog.accept());
		await page.goto(`/?framework=${framework}`);
		const load = async (buffer: Buffer) => {
			await (
				await fileInput(page)
			).setInputFiles({
				name: 'cell-alignment.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer,
			});
		};
		await load(await fixture());
		const editor = page.locator('docx-editor');
		const cells = editor.locator('.ProseMirror td');
		await expect(cells.first()).toHaveCSS('vertical-align', 'bottom');
		await expect(cells.nth(1)).toHaveCSS('vertical-align', 'middle');
		const open = async () => {
			await cells.first().locator('p').click();
			await editor.getByRole('tab', { name: 'Table', exact: true }).click();
			await editor.getByRole('button', { name: 'Table properties', exact: true }).click();
			return editor.getByRole('dialog', { name: 'Table properties', exact: true });
		};
		let dialog = await open();
		const alignment = dialog.getByRole('combobox', {
			name: 'Cell vertical alignment',
			exact: true,
		});
		await expect(alignment).toHaveValue('bottom');
		await alignment.selectOption('center');
		await dialog.getByRole('spinbutton', { name: 'Cell top', exact: true }).fill('0.05');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(cells.first()).toHaveCSS('vertical-align', 'middle');
		await expect(cells.first()).toHaveCSS('padding-top', '4.8px');
		await page.keyboard.press('Control+z');
		await expect(cells.first()).toHaveCSS('vertical-align', 'bottom');
		await expect(cells.first()).toHaveCSS('padding-top', '0px');
		await page.keyboard.press('Control+y');
		await expect(cells.first()).toHaveCSS('vertical-align', 'middle');
		await expect(cells.first()).toHaveCSS('padding-top', '4.8px');
		const save = async () => {
			const pending = page.waitForEvent('download');
			await saveButton(page).click();
			return readFile((await (await pending).path())!);
		};
		const xmlOf = async (bytes: Buffer) =>
			(await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		let bytes = await save();
		let xml = await xmlOf(bytes);
		expect(xml).toContain('<w:vAlign w:val="center" ext:keep="yes"/>');
		expect(xml).toContain('<ext:keep ext:value="yes"/>');
		expect(xml).toContain('<w:tcW w:w="3000" w:type="dxa"/>');
		expect(xml).toContain('w:fill="ABCDEF"');
		expect(xml).toContain('<w:bottom w:val="single" w:sz="4"/>');
		await load(bytes);
		await expect(cells.first()).toHaveCSS('vertical-align', 'middle');
		dialog = await open();
		await dialog.getByRole('combobox', { name: 'Cell vertical alignment' }).selectOption('top');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(cells.first()).toHaveCSS('vertical-align', 'top');
		dialog = await open();
		await dialog.getByRole('combobox', { name: 'Cell vertical alignment' }).selectOption('default');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		bytes = await save();
		xml = await xmlOf(bytes);
		expect(xml.match(/<w:vAlign\b/g)).toHaveLength(1);
		expect(xml).toContain('<w:vAlign w:val="center"/>');
		expect(xml).toContain('<ext:keep ext:value="yes"/>');
		await load(bytes);
		await expect(cells.first()).toHaveCSS('vertical-align', 'top');
		await expect(cells.nth(1)).toHaveCSS('vertical-align', 'middle');
		dialog = await open();
		await expect(dialog.getByRole('combobox', { name: 'Cell vertical alignment' })).toHaveValue(
			'default',
		);
	});
}
