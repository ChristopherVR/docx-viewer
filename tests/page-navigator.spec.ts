import { expect, test, type Page } from '@playwright/test';
import type { DocxEditorElement } from '@christophervr/docx-web-component';
import { openSample } from './helpers';

const editor = (page: Page) => page.locator('docx-editor');

/** Replaces the sample with a long document so Print Layout produces several pages. */
async function useLongDocument(page: Page) {
	await editor(page).evaluate((element) => {
		const host = element as DocxEditorElement;
		const model = structuredClone(host.documentModel!);
		model.blocks = Array.from({ length: 140 }, (_, index) => ({
			type: 'paragraph' as const,
			id: `long-${index}`,
			runs: [{ text: `Paragraph ${index}: pagination in this editor is approximate.` }],
		}));
		host.documentModel = model;
		const events: { page: number; pageCount: number }[] = [];
		host.addEventListener('page-change', (event) => events.push(event.detail));
		(window as unknown as { pageEvents: typeof events }).pageEvents = events;
	});
}

const pageEvents = (page: Page) =>
	page.evaluate(
		() => (window as unknown as { pageEvents: { page: number; pageCount: number }[] }).pageEvents,
	);

test.describe('page thumbnail rail', () => {
	test('toggles from the View tab and explains it needs Print Layout', async ({ page }) => {
		await openSample(page);
		const rail = editor(page).getByRole('complementary', { name: 'Page thumbnails' });
		await expect(rail).toBeHidden();
		await editor(page).getByRole('tab', { name: 'View', exact: true }).click();
		await editor(page).getByRole('button', { name: 'Page thumbnails' }).click();
		await expect(rail).toBeVisible();
		await expect(editor(page).locator('.dve-pages-message')).toContainText('Print Layout');
		await expect(editor(page).locator('.dve-pages-rail').getByRole('option')).toHaveCount(0);
		await editor(page).getByRole('button', { name: 'Close page thumbnails' }).click();
		await expect(rail).toBeHidden();
		await expect(editor(page)).not.toHaveAttribute('show-thumbnails', /.*/);
	});

	test('lists pages, scrolls on click, follows scroll and emits page-change', async ({ page }) => {
		await openSample(page);
		await useLongDocument(page);
		await editor(page).evaluate((element) => {
			(element as DocxEditorElement).showThumbnails = true;
		});
		await editor(page).getByRole('tab', { name: 'View', exact: true }).click();
		await editor(page).getByRole('combobox', { name: 'Layout view' }).selectOption('print');
		const options = editor(page).locator('.dve-pages-rail').getByRole('option');
		await expect.poll(() => options.count()).toBeGreaterThan(2);
		const total = await options.count();
		await expect(options.first()).toHaveAttribute('aria-selected', 'true');
		await expect(editor(page).locator('.dve-status-page')).toHaveText(`Page 1 of ${total}`);
		// Lazy thumbnails: the visible ones are drawn.
		await expect(options.first().locator('.dve-print-page')).toHaveCount(1);

		await options.nth(2).click();
		await expect(options.nth(2)).toHaveAttribute('aria-selected', 'true');
		await expect
			.poll(() =>
				editor(page)
					.locator('.dve-canvas')
					.evaluate((node) => node.scrollTop),
			)
			.toBeGreaterThan(500);
		await expect(editor(page).locator('.dve-status-page')).toHaveText(`Page 3 of ${total}`);
		await expect
			.poll(async () => (await pageEvents(page)).at(-1))
			.toEqual({
				page: 3,
				pageCount: total,
			});

		// Scrolling the document moves the highlight.
		await editor(page)
			.locator('.dve-canvas')
			.evaluate((node) => {
				node.scrollTop = 0;
			});
		await expect(options.first()).toHaveAttribute('aria-selected', 'true');
		await expect.poll(async () => (await pageEvents(page)).at(-1)?.page).toBe(1);
	});

	test('supports arrow-key navigation and Enter', async ({ page }) => {
		await openSample(page);
		await useLongDocument(page);
		await editor(page).evaluate((element) => {
			(element as DocxEditorElement).showThumbnails = true;
		});
		await editor(page).getByRole('tab', { name: 'View', exact: true }).click();
		await editor(page).getByRole('combobox', { name: 'Layout view' }).selectOption('print');
		const options = editor(page).locator('.dve-pages-rail').getByRole('option');
		await expect.poll(() => options.count()).toBeGreaterThan(2);
		await options.first().focus();
		await page.keyboard.press('ArrowDown');
		await expect(options.nth(1)).toBeFocused();
		await page.keyboard.press('ArrowDown');
		await expect(options.nth(2)).toBeFocused();
		await page.keyboard.press('ArrowUp');
		await expect(options.nth(1)).toBeFocused();
		await page.keyboard.press('Enter');
		await expect(options.nth(1)).toHaveAttribute('aria-selected', 'true');
		await expect.poll(async () => (await pageEvents(page)).at(-1)?.page).toBe(2);
	});
});
