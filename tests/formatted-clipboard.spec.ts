import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { newDocument } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: ribbon clipboard preserves formatted paragraphs and a table across stories`, async ({
		page,
	}) => {
		// Exercise the async browser API deterministically without depending on OS clipboard permissions.
		await page.addInitScript(() => {
			let items: ClipboardItem[] = [];
			Object.defineProperty(document, 'execCommand', { value: () => false });
			Object.defineProperty(navigator, 'clipboard', {
				value: {
					read: async () => items,
					write: async (next: ClipboardItem[]) => {
						items = next;
					},
				},
			});
		});
		await page.goto(`/?framework=${framework}`);
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
						endsAtBlockId: model.blocks[0]!.id,
						pageWidthTwips: 12240,
						pageHeightTwips: 15840,
						marginTopTwips: 1440,
						marginBottomTwips: 1440,
						marginLeftTwips: 1440,
						marginRightTwips: 1440,
						columns: { count: 1, equalWidth: true },
						headers: {
							default: {
								partName: 'word/header1.xml',
								blocks: [{ type: 'paragraph', id: 'header', runs: [{ text: 'Header' }] }],
							},
						},
					},
				],
			} as typeof model;
		});
		await page.evaluate(async () =>
			navigator.clipboard.write([
				new ClipboardItem({
					'text/html': new Blob(
						[
							'<p><strong>Bold</strong> text</p><table><tbody><tr><td><p><em>Cell</em></p></td></tr></tbody></table>',
						],
						{ type: 'text/html' },
					),
					'text/plain': new Blob(['Bold text\nCell'], { type: 'text/plain' }),
				}),
			]),
		);
		const body = editor.locator('.dve-paper > .ProseMirror');
		await body.click();
		await editor.getByRole('button', { name: 'Paste', exact: true }).click();
		await expect(body.locator('strong')).toHaveText('Bold');
		await expect(body.locator('table td em')).toHaveText('Cell');
		await body.click();
		await page.keyboard.press('Control+a');
		await editor.getByRole('button', { name: 'Copy', exact: true }).click();
		const header = editor.locator('.dve-header [data-slot=default]');
		await header.dblclick();
		const story = header.locator('.ProseMirror');
		await page.keyboard.press('Control+a');
		await editor.getByRole('tab', { name: 'Home', exact: true }).click();
		await editor.getByRole('button', { name: 'Paste', exact: true }).click();
		await expect(story.locator('strong')).toHaveText('Bold');
		await expect(story.locator('table td em')).toHaveText('Cell');
		await page.keyboard.press('Escape');
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const zip = await JSZip.loadAsync(new Uint8Array(bytes));
		for (const name of ['word/document.xml', 'word/header1.xml']) {
			const xml = await zip.file(name)!.async('string');
			expect(xml).toContain('<w:b');
			expect(xml).toContain('<w:i');
			expect(xml).toContain('<w:tbl');
			expect(xml).toContain('Bold');
			expect(xml).toContain('Cell');
		}
	});
}
