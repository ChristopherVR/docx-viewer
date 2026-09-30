import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { newDocument } from './helpers';

test('scaled glyphs reserve their advances, add spacing after scaling, and wrap at words', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		host.documentModel = {
			...host.documentModel!,
			blocks: [
				...([50, 125, 200, 600, 0] as const).map((scale) => ({
					type: 'paragraph' as const,
					id: `s${scale}`,
					runs: [
						{
							text: 'MMMM MMMM',
							fontFamily: 'Courier New',
							fontSize: 12,
							textScalePercent: scale,
							characterSpacingTwips: 40,
						},
					],
				})),
				{
					type: 'paragraph',
					id: 'wrap',
					runs: [
						{
							text: 'MMMM MMMM MMMM',
							fontFamily: 'Courier New',
							fontSize: 12,
							textScalePercent: 200,
						},
					],
				},
			],
		} as unknown as NonNullable<typeof host.documentModel>;
	});
	await expect(editor.locator('.dve-paper > .ProseMirror p')).toHaveCount(6);
	const metrics = await editor.locator('.dve-paper > .ProseMirror p').evaluateAll((paragraphs) =>
		paragraphs.map((paragraph) => {
			const spans = [...paragraph.querySelectorAll<HTMLElement>('.dve-scaled-text')];
			return spans.map((span) => {
				const css = getComputedStyle(span);
				const canvas = document.createElement('canvas').getContext('2d')!;
				canvas.font = `${css.fontSize} ${css.fontFamily}`;
				return {
					text: span.textContent!,
					left: span.getBoundingClientRect().left,
					top: span.getBoundingClientRect().top,
					width: span.getBoundingClientRect().width,
					glyphWidth: canvas.measureText(span.textContent!).width,
				};
			});
		}),
	);
	for (const [index, scale] of [50, 125, 200, 600, 0].entries()) {
		const [first, space, second] = metrics[index]!;
		const expected = (first!.glyphWidth * scale) / 100 + (4 * 40) / 15;
		if (scale !== 0) expect(first!.width).toBeCloseTo(expected, 0);
		expect(space!.left - first!.left).toBeCloseTo(expected, 0);
		expect(second!.left - space!.left).toBeCloseTo((space!.glyphWidth * scale) / 100 + 40 / 15, 0);
	}
	await editor.locator('.dve-paper > .ProseMirror').evaluate((body) => {
		(body as HTMLElement).style.width = '180px';
	});
	const words = await editor
		.locator('.dve-paper > .ProseMirror p')
		.last()
		.locator('.dve-scaled-text')
		.evaluateAll((spans) =>
			spans
				.filter((span) => span.textContent === 'MMMM')
				.map((span) => ({ top: span.getBoundingClientRect().top })),
		);
	expect(words[1]!.top).toBeCloseTo(words[0]!.top, 1);
	expect(words[2]!.top).toBeGreaterThan(words[0]!.top + 10);
});

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: clicking scaled text places the caret and typing remains one undo step`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			const host = element as DocxEditorElement;
			host.documentModel = {
				...host.documentModel!,
				blocks: [
					{
						type: 'paragraph',
						id: 'scaled',
						runs: [
							{ text: 'MMMM', fontFamily: 'Courier New', fontSize: 12, textScalePercent: 200 },
						],
					},
				],
			};
		});
		const span = editor.locator('.dve-paper > .ProseMirror .dve-scaled-text');
		const box = await span.boundingBox();
		if (!box) throw new Error('Missing scaled text');
		await page.mouse.click(box.x + box.width * 0.5, box.y + box.height * 0.5);
		await page.keyboard.type('X');
		const body = editor.locator('.dve-paper > .ProseMirror');
		await expect(body).toHaveText('MMXMM');
		await page.keyboard.press('Control+z');
		await expect(body).toHaveText('MMMM');
		await page.keyboard.press('Control+y');
		await expect(body).toHaveText('MMXMM');
		const runs = await editor.evaluate((element) => {
			const p = (element as DocxEditorElement).documentModel!.blocks[0]!;
			if (p.type !== 'paragraph') throw new Error('Expected paragraph');
			return p.runs;
		});
		expect(runs.every((run) => run.textScalePercent === 200)).toBe(true);
	});
}

test('formatting boundaries inside a scaled word do not introduce wrapping or document characters', async ({
	page,
}) => {
	test.fail(
		true,
		'Continuous scaled words split across formatting runs still wrap at that boundary; tracked in outstanding-work.md.',
	);
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		host.documentModel = {
			...host.documentModel!,
			blocks: [
				{
					type: 'paragraph',
					id: 'mixed',
					runs: [
						{ text: 'MMMM', fontFamily: 'Courier New', fontSize: 12, textScalePercent: 50 },
						{
							text: 'MMMM',
							fontFamily: 'Courier New',
							fontSize: 12,
							textScalePercent: 50,
							bold: true,
						},
					],
				},
			],
		};
	});
	const body = editor.locator('.dve-paper > .ProseMirror');
	await body.evaluate((root) => {
		(root as HTMLElement).style.width = '30px';
	});
	const rects = await body
		.locator('.dve-scaled-text')
		.evaluateAll((spans) => spans.map((span) => span.getBoundingClientRect().top));
	expect(rects[1]).toBeCloseTo(rects[0]!, 1);
	await expect(body).toHaveText('MMMMMMMM');
	await body.click();
	await page.keyboard.press('Control+a');
	await page.keyboard.type('replacement');
	await expect(body).toHaveText('replacement');
	await page.keyboard.press('Control+z');
	await expect(body).toHaveText('MMMMMMMM');
});
