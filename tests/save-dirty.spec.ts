import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '@christophervr/docx-web-component';
import { openSample } from './helpers';

test.describe('save API and dirty tracking', () => {
	test('dirty follows edits and clears on download and markClean', async ({ page }) => {
		await openSample(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			const flips: boolean[] = [];
			element.addEventListener('dirty-change', (event) =>
				flips.push((event as CustomEvent<boolean>).detail),
			);
			(window as unknown as { flips: boolean[] }).flips = flips;
		});
		expect(await editor.evaluate((element) => (element as DocxEditorElement).dirty)).toBe(false);
		await editor.locator('.ProseMirror').click();
		await page.keyboard.type('Dirty now. ');
		expect(await editor.evaluate((element) => (element as DocxEditorElement).dirty)).toBe(true);
		await expect(editor.locator('.dve-titlebar')).toContainText('Unsaved');

		const download = page.waitForEvent('download');
		await editor.evaluate((element) => (element as DocxEditorElement).download('Edited.docx'));
		expect((await download).suggestedFilename()).toBe('Edited.docx');
		expect(await editor.evaluate((element) => (element as DocxEditorElement).dirty)).toBe(false);

		await page.keyboard.type('More. ');
		expect(await editor.evaluate((element) => (element as DocxEditorElement).dirty)).toBe(true);
		await editor.evaluate((element) => (element as DocxEditorElement).markClean());
		expect(await editor.evaluate((element) => (element as DocxEditorElement).dirty)).toBe(false);
		expect(await page.evaluate(() => (window as unknown as { flips: boolean[] }).flips)).toEqual([
			true,
			false,
			true,
			false,
		]);
	});

	test('File > Save from the title bar shares the same clean state', async ({ page }) => {
		await openSample(page);
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror').click();
		await page.keyboard.type('Changed. ');
		const download = page.waitForEvent('download');
		await editor.getByRole('button', { name: 'Save' }).first().click();
		await download;
		await expect
			.poll(() => editor.evaluate((element) => (element as DocxEditorElement).dirty))
			.toBe(false);
	});

	test('save() returns a docx Blob containing the edit', async ({ page }) => {
		await openSample(page);
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror').click();
		await page.keyboard.type('Saved marker. ');
		const saved = await editor.evaluate(async (element) => {
			const blob = await (element as DocxEditorElement).save();
			return { type: blob.type, bytes: Array.from(new Uint8Array(await blob.arrayBuffer())) };
		});
		expect(saved.type).toContain('wordprocessingml.document');
		const zip = await JSZip.loadAsync(new Uint8Array(saved.bytes));
		expect(await zip.file('word/document.xml')!.async('string')).toContain('Saved marker.');
	});

	test('showToolbar and hiddenActions customise the ribbon', async ({ page }) => {
		await openSample(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			(element as DocxEditorElement).hiddenActions = ['Bold', 'Print'];
		});
		await expect(editor.getByRole('button', { name: 'Bold', exact: true })).toBeHidden();
		await expect(editor.getByRole('button', { name: 'Italic', exact: true })).toBeVisible();
		await editor.evaluate((element) => {
			(element as DocxEditorElement).showToolbar = false;
		});
		await expect(editor.locator('.dve-ribbon')).toBeHidden();
		await expect(editor).toHaveAttribute('show-toolbar', 'false');
		await editor.evaluate((element) => element.setAttribute('show-toolbar', 'true'));
		await expect(editor.locator('.dve-ribbon')).toBeVisible();
	});
});
