import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const frameworks = ['vanilla', 'react', 'vue', 'angular', 'svelte'];
const wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const stylesXml = `<?xml version="1.0"?><w:styles xmlns:w="${wordNs}">
<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:pPr><w:ind w:left="180"/><w:spacing w:after="360"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Base"/><w:pPr><w:spacing w:before="60"/></w:pPr></w:style>
</w:styles>`;
const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${wordNs}"><w:body>
<w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:t>Inherited style sample</w:t></w:r></w:p>
<w:sectPr/></w:body></w:document>`;

async function styledDocx(): Promise<Buffer> {
	const zip = new JSZip();
	zip.file('word/document.xml', documentXml);
	zip.file('word/styles.xml', stylesXml);
	return zip.generateAsync({ type: 'nodebuffer' });
}

for (const framework of frameworks) {
	test(`${framework}: French UI locale stays separate from document metadata and styled DOCX remains unflattened`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.ProseMirror')).toContainText('Document title');
		await page.locator('#file').setInputFiles({
			name: 'styles.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await styledDocx(),
		});
		await expect(page.locator('#filename')).toHaveText('styles.docx');
		const surface = editor.locator('.ProseMirror');
		const styledParagraph = surface.locator('p').filter({ hasText: 'Inherited style sample' });
		await expect(styledParagraph).toHaveCSS('margin-left', '12px');
		await expect(styledParagraph).toHaveCSS('margin-bottom', '24px');
		await expect(editor.getByLabel('Style', { exact: true })).toHaveValue('Derived');

		const modelBefore = await editor.evaluate((element) =>
			JSON.stringify((element as HTMLElement & { documentModel: unknown }).documentModel),
		);
		await editor.evaluate(
			(element) => ((element as HTMLElement & { locale: string }).locale = 'fr'),
		);
		await expect(editor.getByRole('tab', { name: 'Accueil', exact: true })).toBeVisible();
		await expect(editor.getByLabel('Style', { exact: true })).toHaveValue('Derived');
		await expect(
			editor.getByRole('button', { name: 'Rechercher et remplacer', exact: true }),
		).toBeVisible();
		await editor.getByRole('button', { name: 'Rechercher et remplacer', exact: true }).click();
		await expect(editor.getByLabel('Rechercher le texte', { exact: true })).toBeVisible();
		await expect(editor.getByRole('status', { name: 'Résultats de recherche' })).toContainText(
			'Saisissez le texte',
		);
		await page.getByLabel('Read only', { exact: true }).check();
		await expect(
			editor.getByRole('button', { name: 'Tout remplacer', exact: true }),
		).toBeDisabled();
		await page.getByLabel('Read only', { exact: true }).uncheck();

		// Display locale changes labels only; the model, including style catalog and paragraph references, is unchanged.
		const modelAfterFrench = await editor.evaluate((element) =>
			JSON.stringify((element as HTMLElement & { documentModel: unknown }).documentModel),
		);
		expect(modelAfterFrench).toBe(modelBefore);
		await editor.evaluate(
			(element) => ((element as HTMLElement & { locale: string }).locale = 'en'),
		);
		await expect(editor.getByRole('tab', { name: 'Home', exact: true })).toBeVisible();
		await expect(editor.getByLabel('Find text', { exact: true })).toBeVisible();
		const modelAfterEnglish = await editor.evaluate((element) =>
			JSON.stringify((element as HTMLElement & { documentModel: unknown }).documentModel),
		);
		expect(modelAfterEnglish).toBe(modelBefore);

		await styledParagraph.click();
		await page.keyboard.press('End');
		await page.keyboard.type(' edited');
		const pending = page.waitForEvent('download');
		await page.locator('#save').click();
		const download = await pending;
		const zip = await JSZip.loadAsync(await readFile((await download.path())!));
		const savedStyles = await zip.file('word/styles.xml')!.async('string');
		const savedDocument = await zip.file('word/document.xml')!.async('string');
		expect(savedStyles).toBe(stylesXml);
		expect(savedDocument).toContain('<w:pStyle w:val="Derived"');
		expect(savedDocument).not.toContain('w:ind w:left="180"');
		expect(savedDocument).not.toContain('w:spacing w:after="360"');
		await page.locator('#file').setInputFiles({
			name: 'reopened.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await readFile((await download.path())!),
		});
		const reopenedParagraph = editor
			.locator('.ProseMirror p')
			.filter({ hasText: 'Inherited style sample edited' });
		await expect(reopenedParagraph).toHaveCSS('margin-left', '12px');
		await expect(reopenedParagraph).toHaveCSS('margin-bottom', '24px');
		await expect(editor.getByLabel('Style', { exact: true })).toHaveValue('Derived');
		expect(errors).toEqual([]);
	});
}
