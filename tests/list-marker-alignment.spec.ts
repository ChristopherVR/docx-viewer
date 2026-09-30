import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import type { SignedTwips, Twips } from '../packages/core/src/index';
import { newDocument, reveal } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: editor and Print Layout align list markers without adding editable characters`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			const host = element as DocxEditorElement;
			const alignments = ['left', 'center', 'right'] as const;
			host.documentModel = {
				...host.documentModel!,
				numberingCatalog: {
					warnings: [],
					abstractNums: Object.fromEntries(
						alignments.map((alignment, i) => [
							String(i),
							{
								id: String(i),
								levels: {
									0: {
										level: 0,
										start: 12,
										numFmt: 'decimal',
										lvlText: 'Part %1.',
										lvlJc: alignment,
										suffix: 'tab',
								indentLeftTwips: 1440 as SignedTwips,
								hangingTwips: 1080 as Twips,
									},
								},
							},
						]),
					),
					nums: Object.fromEntries(
						alignments.map((_, i) => [
							String(i + 1),
							{ id: String(i + 1), abstractNumId: String(i) },
						]),
					),
				},
				blocks: alignments.map((alignment, i) => ({
					type: 'paragraph',
					id: alignment,
					runs: [{ text: 'Body text', fontFamily: 'Courier New', fontSize: 12 }],
					numbering: { numId: i + 1, level: 0 },
				})),
			};
		});
		await page.evaluate(() => document.fonts.ready);
		const paragraphs = editor.locator('.dve-paper > .ProseMirror p');
		await expect(paragraphs).toHaveText(['Body text', 'Body text', 'Body text']);
		const geometry = await paragraphs.evaluateAll((nodes) =>
			nodes.map((node) => {
				const style = getComputedStyle(node),
					marker = getComputedStyle(node, '::before');
				const range = document.createRange();
				const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT);
				const text = walker.nextNode()!;
				range.setStart(text, 0);
				range.setEnd(text, 1);
				const width = parseFloat(marker.width);
				return {
					marker: parseFloat(style.marginLeft) + parseFloat(marker.left),
					width,
					body:
						range.getBoundingClientRect().left -
						node.getBoundingClientRect().left +
						parseFloat(style.marginLeft),
				};
			}),
		);
		for (const [i, factor] of [0, 0.5, 1].entries())
			expect(geometry[i]!.marker + factor * geometry[i]!.width).toBeCloseTo(24, 1);
		expect(geometry[1]!.body).toBeCloseTo(96, 1);
		expect(geometry[2]!.body).toBeCloseTo(96, 1);
		await paragraphs.nth(2).click();
		await page.keyboard.press('Home');
		await page.keyboard.type('X');
		await expect(paragraphs.nth(2)).toHaveText('XBody text');
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		const print = editor.getByRole('button', { name: 'Print Layout', exact: true });
		await reveal(editor, print);
		await print.click();
		const lines = editor.locator('.dve-print-line');
		await expect(lines).toHaveCount(3);
		const positions = await lines.evaluateAll((nodes) =>
			nodes.map((node) =>
				[...node.querySelectorAll('span')]
					.filter((span) => span.textContent)
					.map((span) => ({
						text: span.textContent,
						x: parseFloat(span.style.left),
						width: span.getBoundingClientRect().width,
					})),
			),
		);
		for (const [i, factor] of [0, 0.5, 1].entries()) {
			const marker = positions[i]![0]!;
			expect(marker.text).toBe('Part 12.');
			expect(marker.x + factor * marker.width).toBeCloseTo(24, 1);
			expect(marker.x).toBeCloseTo(geometry[i]!.marker, 1);
		}
	});
}
