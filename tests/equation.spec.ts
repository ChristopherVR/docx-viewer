import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import { equationFixture, fractionOmml, displayOmml } from './support/equation-fixture';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: inline and display equations render and survive text edits and reopening`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'equations.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(await equationFixture()),
		});
		const editor = page.locator('docx-editor');
		const equations = editor.locator('.ProseMirror [data-docx-equation]');
		await expect(equations).toHaveCount(2);
		await expect(equations.nth(0).locator('mfrac')).toHaveCount(1);
		await expect(equations.nth(1).locator('msup')).toHaveCount(1);
		await expect(equations.nth(1).locator('math')).toHaveAttribute('display', 'block');
		for (const equation of await equations.all()) {
			await expect(equation).toHaveAttribute('contenteditable', 'false');
			await expect(equation.locator('math')).toHaveAttribute('aria-label', /\S/);
			expect(
				await equation.locator('math').evaluate((math) => math.getBoundingClientRect().height),
			).toBeGreaterThan(10);
		}
		const inlineParagraph = editor.locator('.dve-paper > .ProseMirror > p').first();
		await inlineParagraph.click();
		await page.keyboard.press('Home');
		await page.keyboard.type('Edited ');
		const paragraph = editor.locator('.dve-paper > .ProseMirror > p').last();
		await paragraph.click();
		await page.keyboard.press('End');
		await page.keyboard.type(' changed');
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const zip = await JSZip.loadAsync(new Uint8Array(bytes));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain(fractionOmml);
		expect(xml).toContain(displayOmml);
		expect(xml).toContain('Editable paragraph changed');
		expect(xml).toContain('Edited Before ');
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'saved-equations.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		await expect(equations).toHaveCount(2);
		await expect(equations.nth(0).locator('mfrac')).toHaveCount(1);
		await expect(equations.nth(1).locator('msup')).toHaveCount(1);
		if (framework === 'vanilla') {
			await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
			await expect(editor.locator('.dve-print-page').first()).toContainText('[Equation]');
		}
		expect(errors).toEqual([]);
	});
}
