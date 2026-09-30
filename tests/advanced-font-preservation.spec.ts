import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { loadDocx } from '../packages/core/src/index';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { newDocument, saveButton } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: editing and exporting keeps advanced font settings`, async ({ page }) => {
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
						id: 'advanced',
						runs: [
							{
								text: 'Advanced',
								fontFamily: 'Courier New',
								fontSize: 12,
								textScalePercent: 125,
								kerningHalfPoints: 24,
								positionHalfPoints: -6,
							},
						],
					},
				],
			} as NonNullable<typeof host.documentModel>;
		});
		const body = editor.locator('.ProseMirror').first();
		await body.click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' changed');
		await expect(body).toHaveText('Advanced changed');
		await editor
			.locator('.dve-quick-access')
			.getByRole('button', { name: 'Undo', exact: true })
			.click();
		await expect(body).toHaveText('Advanced');
		await editor
			.locator('.dve-quick-access')
			.getByRole('button', { name: 'Redo', exact: true })
			.click();
		await expect(body).toHaveText('Advanced changed');
		const download = page.waitForEvent('download');
		await saveButton(page).click();
		const artifact = await download;
		const path = await artifact.path();
		if (!path) throw new Error('Expected local DOCX download');
		const { model } = await loadDocx(await readFile(path));
		const paragraph = model.blocks[0];
		if (paragraph?.type !== 'paragraph') throw new Error('Expected exported paragraph');
		expect(paragraph.runs.map((run) => run.text).join('')).toBe('Advanced changed');
		for (const run of paragraph.runs)
			expect(run).toMatchObject({
				textScalePercent: 125,
				kerningHalfPoints: 24,
				positionHalfPoints: -6,
			});
		await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
		const fragment = editor
			.locator('.dve-print-line span')
			.filter({ hasText: /^Advanced$/ })
			.first();
		await expect(fragment).toBeVisible();
		const metrics = await fragment.evaluate((element) => {
			const span = element as HTMLElement;
			const css = getComputedStyle(span);
			const canvas = document.createElement('canvas').getContext('2d')!;
			canvas.font = `${css.fontSize} ${css.fontFamily}`;
			canvas.fontKerning = 'normal';
			return {
				actual: span.getBoundingClientRect().width,
				expected: canvas.measureText(span.textContent!).width * 1.25,
				kerning: css.fontKerning,
				transform: span.style.transform,
			};
		});
		expect(metrics.actual).toBeCloseTo(metrics.expected, 0);
		expect(metrics.kerning).toBe('normal');
		expect(metrics.transform).toBe('scaleX(1.25)');
	});
}

test('Print Layout keeps spacing independent of scale and aligns raised/lowered baselines', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		const font = { fontFamily: 'Courier New', fontSize: 12 };
		host.documentModel = {
			...host.documentModel!,
			blocks: [
				{
					type: 'paragraph',
					id: 'spacing',
					runs: [{ text: 'MMMMMMMMMM', ...font, textScalePercent: 200, characterSpacingTwips: 40 }],
				},
				{
					type: 'paragraph',
					id: 'position',
					runs: [
						{ text: 'Normal', ...font },
						{ text: 'Raised', ...font, positionHalfPoints: 6 },
						{ text: 'Lowered', ...font, positionHalfPoints: -6 },
					],
				},
			],
		} as NonNullable<typeof host.documentModel>;
	});
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const spans = editor.locator('.dve-print-line span');
	const scaled = spans.filter({ hasText: /^MMMMMMMMMM$/ });
	await expect(scaled).toBeVisible();
	const widths = await scaled.evaluate((element) => {
		const css = getComputedStyle(element);
		const canvas = document.createElement('canvas').getContext('2d')!;
		canvas.font = `${css.fontSize} ${css.fontFamily}`;
		return {
			actual: element.getBoundingClientRect().width,
			expected: canvas.measureText(element.textContent!).width * 2 + (10 * 40) / 15,
		};
	});
	expect(widths.actual).toBeCloseTo(widths.expected, 0);
	const tops = await Promise.all(
		['Normal', 'Raised', 'Lowered'].map((text) =>
			spans
				.filter({ hasText: new RegExp(`^${text}$`) })
				.evaluate((element) => element.getBoundingClientRect().top),
		),
	);
	expect(tops[0]! - tops[1]!).toBeCloseTo(4, 1);
	expect(tops[2]! - tops[0]!).toBeCloseTo(4, 1);
});
