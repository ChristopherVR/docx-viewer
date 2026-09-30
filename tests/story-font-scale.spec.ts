import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { newDocument } from './helpers';

test('inherited scale renders in header/footer/note previews, editors and after closing without flattening runs', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		const paragraph = (id: string) => ({
			type: 'paragraph' as const,
			id,
			runs: [{ text: 'Story text' }],
		});
		host.documentModel = {
			...host.documentModel!,
			characterStyles: {
				docDefaults: { fontFamily: 'DVE Scale Story', fontSize: 12, textScalePercent: 150 },
				styles: {},
				warnings: [],
			},
			blocks: [
				{
					type: 'paragraph',
					id: 'body',
					runs: [
						{ text: 'Body', textScalePercent: 100 },
						{ text: '', noteReference: { kind: 'footnote', id: '1' } },
					],
				},
			],
			sections: [
				{
					type: 'nextPage',
					orientation: 'portrait',
					pageWidthTwips: 12240,
					pageHeightTwips: 15840,
					marginTopTwips: 1440,
					marginBottomTwips: 1440,
					marginLeftTwips: 1440,
					marginRightTwips: 1440,
					columns: { count: 1, equalWidth: true },
					endsAtBlockId: 'body',
					headers: { default: { partName: 'word/header1.xml', blocks: [paragraph('header')] } },
					footers: { default: { partName: 'word/footer1.xml', blocks: [paragraph('footer')] } },
				},
			],
			footnotes: [{ id: '1', blocks: [paragraph('note')] }],
		} as unknown as NonNullable<typeof host.documentModel>;
	});
	await expect(editor.locator('.dve-paper > .ProseMirror .dve-scaled-text')).toHaveCount(0);
	const header = editor.locator('.dve-header .dve-scaled-text');
	const advance = () =>
		header.evaluateAll((spans) => ({
			width: spans[0]!.getBoundingClientRect().width,
			advance: spans[1]!.getBoundingClientRect().left - spans[0]!.getBoundingClientRect().left,
		}));
	const beforeFont = await advance();
	expect(beforeFont.advance).toBeCloseTo(beforeFont.width, 0);
	await page.evaluate(async () => {
		const font = new FontFace('DVE Scale Story', 'local("Courier New")');
		await font.load();
		document.fonts.add(font);
		document.fonts.dispatchEvent(new Event('loadingdone'));
	});
	const afterFont = await advance();
	expect(afterFont.advance).toBeCloseTo(afterFont.width, 0);
	expect(Math.abs(afterFont.width - beforeFont.width)).toBeGreaterThan(1);
	for (const selector of [
		'.dve-header [data-slot=default]',
		'.dve-footer [data-slot=default]',
		'.dve-note-list li',
	]) {
		const story = editor.locator(selector);
		const read = () =>
			story
				.locator('.dve-scaled-text')
				.first()
				.evaluate((span) => {
					const css = getComputedStyle(span);
					const canvas = document.createElement('canvas').getContext('2d')!;
					canvas.font = `${css.fontSize} ${css.fontFamily}`;
					return {
						width: span.getBoundingClientRect().width,
						expected: canvas.measureText(span.textContent!).width * 1.5,
						text: span.textContent,
						transform: css.transform,
					};
				});
		const preview = await read();
		expect(preview.width).toBeCloseTo(preview.expected, 0);
		await story.dblclick();
		const editing = await read();
		expect(editing.width).toBeCloseTo(preview.width, 1);
		await page.keyboard.press('Escape');
		expect((await read()).width).toBeCloseTo(preview.width, 1);
	}
	const scales = await editor.evaluate((element) => {
		const model = (element as DocxEditorElement).documentModel!;
		return [
			model.sections![0]!.headers!.default!.blocks[0],
			model.sections![0]!.footers!.default!.blocks[0],
			model.footnotes![0]!.blocks[0],
		].map((block) => (block?.type === 'paragraph' ? block.runs[0]!.textScalePercent : null));
	});
	expect(scales).toEqual([undefined, undefined, undefined]);
});
