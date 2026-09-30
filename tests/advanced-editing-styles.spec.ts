import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../packages/web-component/src/component';
import { newDocument } from './helpers';
import { fileInput } from './helpers';
import JSZip from 'jszip';

test('direct and inherited scripts keep their smaller size and compose with baseline position', async ({
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
				docDefaults: { fontSize: 12 },
				styles: {
					Script: { id: 'Script', type: 'paragraph', formatting: { verticalAlign: 'superscript' } },
					Raised: { id: 'Raised', type: 'paragraph', formatting: { positionHalfPoints: 6 } },
				},
				warnings: [],
			},
			paragraphStyles: {
				styles: {
					Script: { id: 'Script', name: 'Script' },
					Raised: {
						id: 'Raised',
						name: 'Raised',
						basedOn: 'Script',
					},
				},
				warnings: [],
			},
			blocks: [
				{
					type: 'paragraph',
					id: 'direct',
					runs: [
						{ text: 'Normal' },
						{ text: 'Super', verticalAlign: 'superscript' },
						{ text: 'Sub', verticalAlign: 'subscript' },
					],
				},
				{ type: 'paragraph', id: 'inherited', style: 'Script', runs: [{ text: 'Inherited' }] },
				{ type: 'paragraph', id: 'raised', style: 'Raised', runs: [{ text: 'Raised' }] },
			],
		} as unknown as NonNullable<typeof host.documentModel>;
	});
	const metrics = await editor.locator('.dve-paper > .ProseMirror').evaluate((body) => {
		const walker = document.createTreeWalker(body, NodeFilter.SHOW_TEXT);
		const result: Record<string, { size: number; align: string; top: number }> = {};
		let node: Node | null;
		while ((node = walker.nextNode())) {
			const css = getComputedStyle(node.parentElement!);
			result[node.textContent!] = {
				size: parseFloat(css.fontSize),
				align: css.verticalAlign,
				top: parseFloat(css.top),
			};
		}
		return result;
	});
	for (const text of ['Super', 'Sub', 'Inherited', 'Raised'])
		expect(metrics[text]!.size / metrics.Normal!.size).toBeCloseTo(0.65, 4);
	expect(metrics.Inherited!.align).toBe('super');
	expect(metrics.Raised!.align).toBe('super');
	expect(metrics.Raised!.top).toBeCloseTo(-4, 4);
});

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: direct and inherited baseline, spacing and kerning render while editing`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			const host = element as DocxEditorElement;
			host.documentModel = {
				...host.documentModel!,
				characterStyles: {
					docDefaults: {
						fontFamily: 'Times New Roman',
						fontSize: 12,
						positionHalfPoints: 6,
						characterSpacingTwips: 40,
						kerningHalfPoints: 24,
					},
					styles: {},
					warnings: [],
				},
				blocks: [
					{
						type: 'paragraph',
						id: 'advanced',
						runs: [
							{
								text: 'Normal',
								positionHalfPoints: 0,
								characterSpacingTwips: 0,
								kerningHalfPoints: 0,
							},
							{ text: 'Raised', positionHalfPoints: 6, characterSpacingTwips: 0 },
							{
								text: 'Lowered',
								positionHalfPoints: -6,
								characterSpacingTwips: 0,
								kerningHalfPoints: 0,
							},
							{ text: 'Inherited' },
							{ text: 'Small', fontSize: 10, positionHalfPoints: 0, characterSpacingTwips: 0 },
						],
					},
				],
			} as unknown as NonNullable<typeof host.documentModel>;
		});
		const metrics = await editor.locator('.dve-paper > .ProseMirror p').evaluate((paragraph) => {
			const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
			const values: Record<string, { top: number; kerning: string; spacing: string }> = {};
			let node: Node | null;
			while ((node = walker.nextNode())) {
				const range = document.createRange();
				range.selectNodeContents(node);
				const css = getComputedStyle(node.parentElement!);
				values[node.textContent!] = {
					top: range.getBoundingClientRect().top,
					kerning: css.fontKerning,
					spacing: css.letterSpacing,
				};
			}
			return values;
		});
		expect(metrics.Normal!.top - metrics.Raised!.top).toBeCloseTo(4, 1);
		expect(metrics.Lowered!.top - metrics.Normal!.top).toBeCloseTo(4, 1);
		expect(metrics.Normal!.top - metrics.Inherited!.top).toBeCloseTo(4, 1);
		expect(metrics.Raised!.kerning).toBe('normal');
		expect(metrics.Normal!.kerning).toBe('none');
		expect(metrics.Small!.kerning).toBe('none');
		expect(parseFloat(metrics.Inherited!.spacing)).toBeCloseTo(40 / 15, 4);
	});
}

test('an imported paragraph without pStyle uses Normal font settings in the surface and toolbar', async ({
	page,
}) => {
	const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:t>Implicit Normal</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/styles.xml',
		`<w:styles xmlns:w="${ns}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="32"/><w:position w:val="6"/><w:spacing w:val="40"/><w:kern w:val="24"/></w:rPr></w:style></w:styles>`,
	);
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'implicit-normal.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await zip.generateAsync({ type: 'nodebuffer' }),
	});
	const editor = page.locator('docx-editor');
	const text = editor.locator('.dve-paper > .ProseMirror p span').first();
	await expect(text).toHaveText('Implicit Normal');
	const metrics = await text.evaluate((element) => {
		const css = getComputedStyle(element);
		return { size: css.fontSize, kerning: css.fontKerning, position: css.verticalAlign };
	});
	expect(parseFloat(metrics.size)).toBeCloseTo((16 * 4) / 3, 3);
	expect(metrics.kerning).toBe('normal');
	expect(parseFloat(metrics.position)).toBeCloseTo(4, 3);
	await editor.locator('.dve-paper > .ProseMirror').click();
	await expect(editor.getByRole('combobox', { name: 'Font size', exact: true })).toHaveValue('16');
});

test('header and note previews retain inherited formatting while opening and closing their editors', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		const model = host.documentModel!;
		host.documentModel = {
			...model,
			characterStyles: {
				docDefaults: {
					fontFamily: 'Times New Roman',
					fontSize: 12,
					positionHalfPoints: 6,
					characterSpacingTwips: 40,
					kerningHalfPoints: 24,
				},
				styles: {},
				warnings: [],
			},
			blocks: [
				{
					type: 'paragraph',
					id: 'body',
					runs: [{ text: 'Body' }, { text: '', noteReference: { kind: 'footnote', id: '1' } }],
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
					headers: {
						default: {
							partName: 'word/header1.xml',
							blocks: [{ type: 'paragraph', id: 'header', runs: [{ text: 'Header' }] }],
						},
					},
				},
			],
			footnotes: [
				{ id: '1', blocks: [{ type: 'paragraph', id: 'note', runs: [{ text: 'Note' }] }] },
			],
		} as unknown as typeof model;
	});
	for (const [container, text] of [
		['.dve-header [data-slot=default]', 'Header'],
		['.dve-note-list li', 'Note'],
	] as const) {
		const story = editor.locator(container);
		const read = () =>
			story.locator('p').evaluate((paragraph, label) => {
				const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
				let node: Node | null;
				while ((node = walker.nextNode()))
					if (node.textContent === label) {
						const css = getComputedStyle(node.parentElement!);
						return {
							position: css.verticalAlign,
							spacing: css.letterSpacing,
							kerning: css.fontKerning,
							size: css.fontSize,
						};
					}
				throw new Error('Missing story text');
			}, text);
		const preview = await read();
		expect(parseFloat(preview.position)).toBeCloseTo(4, 4);
		expect(preview.kerning).toBe('normal');
		expect(parseFloat(preview.spacing)).toBeCloseTo(40 / 15, 4);
		await story.dblclick();
		expect(await read()).toEqual(preview);
		await page.keyboard.press('Escape');
		expect(await read()).toEqual(preview);
	}
	const source = await editor.evaluate((element) => {
		const model = (element as DocxEditorElement).documentModel!;
		const header = model.sections![0]!.headers!.default!.blocks[0]!;
		const note = model.footnotes![0]!.blocks[0]!;
		if (header.type !== 'paragraph' || note.type !== 'paragraph')
			throw new Error('Expected paragraphs');
		return [header.runs[0]!.positionHalfPoints, note.runs[0]!.positionHalfPoints];
	});
	expect(source).toEqual([undefined, undefined]);
});
