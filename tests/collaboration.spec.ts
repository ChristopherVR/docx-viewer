import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src';
import JSZip from 'jszip';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: coauthors converge after paused concurrent edits`, async ({ page }) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/collaboration.html?framework=${framework}`);
		const a = page.locator('#peer-a docx-editor');
		const b = page.locator('#peer-b docx-editor');
		const textA = a.locator('.ProseMirror');
		const textB = b.locator('.ProseMirror');
		await expect(textA).toContainText('Shared document');
		await expect(textB).toContainText('Shared document');
		await expect(page.locator('#collaboration-status')).toContainText('Synced');
		await page.getByRole('button', { name: 'Pause delivery', exact: true }).click();
		await textA.click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' Alice');
		await page.keyboard.press('Enter');
		await page.keyboard.type('Alice paragraph');
		await textB.click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' Bob');
		await page.keyboard.press('Enter');
		await page.keyboard.type('Bob paragraph');
		await expect(textA).not.toContainText('Bob');
		await expect(textB).not.toContainText('Alice');
		await page.getByRole('button', { name: 'Resume delivery', exact: true }).click();
		await expect(textA).toContainText('Bob paragraph');
		await expect(textB).toContainText('Alice paragraph');
		await expect
			.poll(async () => {
				const models = await Promise.all(
					[a, b].map((peer) =>
						peer.evaluate((element) =>
							JSON.stringify((element as DocxEditorElement).documentModel),
						),
					),
				);
				return models[0] === models[1];
			})
			.toBe(true);
		await expect(page.locator('#collaboration-status')).toContainText('Synced');
		await page.getByLabel('Editor B read only', { exact: true }).check();
		await expect(textB).toHaveAttribute('contenteditable', 'false');
		await textA.click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' Remote updates still arrive.');
		await expect(textB).toContainText('Remote updates still arrive.');
		const ids = await a.evaluate((element) =>
			(element as DocxEditorElement).documentModel!.blocks.map((block) => block.id),
		);
		expect(new Set(ids).size).toBe(ids.length);
		const bytes = await b.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).save()),
		);
		const zip = await JSZip.loadAsync(new Uint8Array(bytes));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('Alice paragraph');
		expect(xml).toContain('Bob paragraph');
		expect(errors).toEqual([]);
	});
}
