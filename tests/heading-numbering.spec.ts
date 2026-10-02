import { expect, test } from '@playwright/test';
import { headingLabels, headingNumberingFixture } from './support/heading-numbering-fixture';
import { fileInput } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: heading-linked outline markers retain their levels after editing`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'headings.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(await headingNumberingFixture()),
		});
		const paragraphs = page.locator('docx-editor .dve-paper > .ProseMirror p[data-list-label]');
		await expect(paragraphs).toHaveCount(5);
		const labels = () =>
			paragraphs.evaluateAll((nodes) =>
				nodes.map((node) => node.getAttribute('data-list-label')?.trim()),
			);
		await expect.poll(labels).toEqual(headingLabels);
		await page.evaluate(() => document.fonts.ready);
		await paragraphs.nth(1).click();
		await page.keyboard.press('Home');
		await page.keyboard.type('Edited ');
		await expect(paragraphs.nth(1)).toHaveText('Edited Heading 1');
		await page
			.getByRole('combobox', { name: 'Multilevel list', exact: true })
			.selectOption('define');
		const dialog = page.getByRole('dialog', { name: 'Define New Multilevel List', exact: true });
		await expect(
			dialog.getByRole('combobox', { name: 'Level to modify', exact: true }),
		).toHaveValue('1');
		await expect(
			dialog.getByRole('textbox', { name: 'Enter formatting for number', exact: true }),
		).toHaveValue('%1.%2.');
		await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
		await expect.poll(labels).toEqual(headingLabels);
		await expect.poll(labels).toEqual(headingLabels);
	});
}
