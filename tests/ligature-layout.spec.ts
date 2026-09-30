import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { newDocument, reveal } from './helpers';

test('ligatures render inherited and explicit-off features and use matching Print Layout advances', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		host.documentModel = {
			...host.documentModel!,
			characterStyles: {
				docDefaults: { fontFamily: 'Georgia', fontSize: 24, ligatures: 'all' },
				styles: {},
				warnings: [],
			},
			blocks: [
				{
					type: 'paragraph',
					id: 'ligatures',
					runs: [
						{ text: 'office' },
						{ text: 'affinity', ligatures: 'none' },
						{ text: 'finish', ligatures: 'standardContextual' },
					],
				},
			],
		};
	});
	const body = editor.locator('.dve-paper > .ProseMirror');
	const settings = await body.evaluate((root) => {
		const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
		const result: Record<string, string> = {};
		let node: Node | null;
		while ((node = walker.nextNode()))
			result[node.textContent!] = getComputedStyle(node.parentElement!).fontVariantLigatures;
		return result;
	});
	expect(settings.office).toContain('historical-ligatures');
	expect(settings.affinity).toBe('none');
	expect(settings.finish).toContain('common-ligatures');
	const viewTab = editor.getByRole('tab', { name: 'View', exact: true });
	await viewTab.click();
	const print = editor.getByRole('button', { name: 'Print Layout', exact: true });
	await reveal(editor, print);
	await print.click();
	const spans = editor.locator('.dve-print-line span');
	await expect(spans.first()).toHaveText('office');
	const positions = await spans.evaluateAll((elements) =>
		elements.map((element) => ({
			text: element.textContent!,
			left: parseFloat((element as HTMLElement).style.left),
			width: element.getBoundingClientRect().width,
			settings: getComputedStyle(element).fontVariantLigatures,
		})),
	);
	for (let index = 0; index < positions.length - 1; index++)
		expect(positions[index + 1]!.left - positions[index]!.left).toBeCloseTo(
			positions[index]!.width,
			1,
		);
	expect(positions.map((value) => value.settings)).toEqual([
		settings.office,
		settings.affinity,
		settings.finish,
	]);
	expect(await page.locator('body > span[style*="100000"]').count()).toBe(0);
});
