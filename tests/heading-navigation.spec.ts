import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { openSample, reveal } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'svelte', 'angular', 'jquery']) {
	test(`heading navigation selects, refreshes and works read-only (${framework})`, async ({
		page,
	}) => {
		await openSample(page, framework);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			const host = element as DocxEditorElement;
			host.documentModel = {
				...host.documentModel!,
				blocks: [
					{ type: 'paragraph', id: 'chapter', style: 'Heading1', runs: [{ text: 'Chapter' }] },
					{ type: 'paragraph', id: 'topic', style: 'Heading2', runs: [{ text: 'Topic' }] },
					{ type: 'paragraph', id: 'body', runs: [{ text: 'Body' }] },
					{ type: 'paragraph', id: 'next', style: 'Heading1', runs: [{ text: 'Next chapter' }] },
				],
			};
		});
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		const button = editor.getByRole('button', { name: 'Navigation pane', exact: true });
		await reveal(editor, button);
		await button.click();
		await expect(button).toHaveAttribute('aria-pressed', 'true');
		const rail = editor.getByRole('complementary', { name: 'Navigation pane', exact: true });
		const headings = rail.getByRole('treeitem');
		await expect(headings.first()).toBeVisible();
		const label = await headings.first().locator('.dve-heading-text').textContent();
		await headings.first().click();
		await page.keyboard.press('Home');
		await page.keyboard.type('Updated ');
		await expect(headings.first()).toContainText(`Updated ${label}`);
		await page.keyboard.press('Control+z');
		await expect(headings.first()).toContainText(label!);
		await editor.evaluate((element) => {
			(element as DocxEditorElement).readOnly = true;
		});
		await expect(button).toBeEnabled();
		await headings.first().focus();
		await page.keyboard.press('End');
		await page.keyboard.press('Enter');
		await expect(headings.last()).toHaveAttribute('aria-selected', 'true');
		await editor.evaluate((element) => element.setAttribute('locale', 'fr'));
		await expect(
			editor.getByRole('complementary', { name: 'Volet de navigation', exact: true }),
		).toBeVisible();
		await editor
			.getByRole('button', { name: 'Fermer le volet de navigation', exact: true })
			.click();
		await expect(rail).toBeHidden();
	});
}
