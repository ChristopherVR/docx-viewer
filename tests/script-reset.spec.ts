import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { loadDocx } from '../packages/core/src/index';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { fileInput, reveal, saveButton } from './helpers';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const styles = `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>`;
async function inheritedDoc() {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${officeRel}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${officeRel}/styles" Target="styles.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>Inherited</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file('word/styles.xml', styles);
	return zip.generateAsync({ type: 'nodebuffer' });
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: ribbon and Font dialog reset inherited script, undo and export baseline`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'inherited-script.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await inheritedDoc(),
		});
		const editor = page.locator('docx-editor');
		const body = editor.locator('.dve-paper > .ProseMirror');
		await expect(body).toHaveText('Inherited');
		await body.click();
		await page.keyboard.press('Control+a');
		const superscript = editor.getByRole('button', {
			name: 'Superscript',
			exact: true,
		});
		await reveal(editor, superscript);
		await expect(superscript).toHaveAttribute('aria-pressed', 'true');
		const size = () =>
			body.locator('p').evaluate((p) => {
				const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
				return parseFloat(getComputedStyle(walker.nextNode()!.parentElement!).fontSize);
			});
		const scriptSize = await size();
		await superscript.click();
		await expect(superscript).toHaveAttribute('aria-pressed', 'false');
		expect(await size()).toBeCloseTo(scriptSize / 0.65, 3);
		await editor
			.locator('.dve-quick-access')
			.getByRole('button', { name: 'Undo', exact: true })
			.click();
		await expect(superscript).toHaveAttribute('aria-pressed', 'true');
		await editor
			.locator('.dve-quick-access')
			.getByRole('button', { name: 'Redo', exact: true })
			.click();
		const launcher = editor.getByRole('button', { name: 'Font settings', exact: true });
		await reveal(editor, launcher);
		await launcher.click();
		const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
		await expect(
			dialog.getByRole('checkbox', { name: 'Superscript', exact: true }),
		).not.toBeChecked();
		await dialog.getByRole('checkbox', { name: 'Subscript', exact: true }).check();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await launcher.click();
		await dialog.getByRole('checkbox', { name: 'Subscript', exact: true }).uncheck();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		const run = await editor.evaluate((element) => {
			const p = (element as DocxEditorElement).documentModel!.blocks[0]!;
			if (p.type !== 'paragraph') throw new Error('Expected paragraph');
			return p.runs[0]!;
		});
		expect(run.verticalAlign).toBe('baseline');
		expect(await size()).toBeCloseTo(16, 3);
		const download = page.waitForEvent('download');
		await saveButton(page).click();
		const path = test.info().outputPath('script-reset.docx');
		await (await download).saveAs(path);
		const bytes = await readFile(path);
		const zip = await JSZip.loadAsync(bytes);
		expect(await zip.file('word/styles.xml')!.async('string')).toBe(styles);
		const p = (await loadDocx(bytes)).model.blocks[0]!;
		if (p.type !== 'paragraph') throw new Error('Expected exported paragraph');
		expect(p.runs[0]!.verticalAlign).toBe('baseline');
	});
}
