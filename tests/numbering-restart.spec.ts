import { expect, test } from '@playwright/test';
import { restartCases, restartFixture } from './support/restart-fixture';
import { fileInput } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: imported multilevel restart rules drive displayed markers`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		for (const { name, restart, expected } of restartCases) {
			await (
				await fileInput(page)
			).setInputFiles({
				name: `${name}.docx`,
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer: Buffer.from(await restartFixture(restart)),
			});
			const paragraphs = page.locator('docx-editor .dve-paper > .ProseMirror p[data-list-label]');
			await expect(paragraphs).toHaveCount(8);
			await expect
				.poll(() =>
					paragraphs.evaluateAll((nodes) =>
						nodes.map((node) => node.getAttribute('data-list-label')?.trim()),
					),
				)
				.toEqual(expected.map((n) => `${n}.`));
		}
	});
}
