import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: Print Layout clicks place the cursor at visible text boundaries`, async ({
		page,
	}) => {
		page.on('dialog', (dialog) => void dialog.accept());
		await page.goto(`/?framework=${framework}`);
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>Wi</w:t></w:r><w:r><w:rPr><w:w w:val="150"/></w:rPr><w:t>de</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
		);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'print-cursor.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await zip.generateAsync({ type: 'nodebuffer' }),
		});
		const editor = page.locator('docx-editor');
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		const print = editor.getByRole('button', { name: 'Print Layout', exact: true });
		await print.click();
		const finalFragment = editor.locator('.dve-print-line span').filter({ hasText: 'de' }).first();
		await expect(finalFragment).toBeVisible();
		await finalFragment.click({
			position: { x: (await finalFragment.boundingBox())!.width - 1, y: 5 },
		});
		await page.keyboard.type('X');
		await expect(editor.locator('.ProseMirror p').first()).toHaveText('WideX');
		await page.keyboard.press('Control+z');
		await print.click();
		const firstFragment = editor.locator('.dve-print-line span').filter({ hasText: 'Wi' }).first();
		await firstFragment.click({ position: { x: 1, y: 5 } });
		await page.keyboard.type('Y');
		await expect(editor.locator('.ProseMirror p').first()).toHaveText('YWide');
	});
}
