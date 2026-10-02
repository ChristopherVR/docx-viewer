// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel, type Paragraph, type WatermarkSpec } from 'docx-core';
import { createWatermarkDialog } from './watermark-dialog';
import { currentWatermark, withWatermark } from './watermark-commands';
import { watermarkElement } from './print-header-footer';

const spec: WatermarkSpec = {
	text: 'DRAFT',
	color: '#c0c0c0',
	semitransparent: true,
	layout: 'diagonal',
	fontFamily: 'Calibri',
};
let counter = 0;
const newId = () => `w${++counter}`;
const headerRuns = (model: DocumentModel, section: number, slot: 'default' | 'first' | 'even') =>
	(model.sections![section]!.headers![slot]!.blocks[0] as Paragraph).runs;

describe('withWatermark', () => {
	it('creates a default header for a document that has none and adds the watermark run first', () => {
		const next = withWatermark(createDocument(), spec, newId);
		const runs = headerRuns(next, 0, 'default');
		expect(runs).toHaveLength(1);
		expect(runs[0]!.image).toMatchObject({ unsupported: 'Watermark', watermark: spec });
		expect(currentWatermark(next)).toEqual(spec);
	});

	it('updates every existing header slot once per part and removes cleanly', () => {
		const base = withWatermark(createDocument(), undefined, newId);
		const blank = withWatermark(
			{
				...base,
				sections: [
					{
						...(base.sections ?? [])[0]!,
						titlePage: true,
						headers: {
							default: {
								partName: 'word/header1.xml',
								blocks: [{ type: 'paragraph', id: 'a', runs: [{ text: 'Head' }] }],
							},
							first: {
								partName: 'word/header2.xml',
								blocks: [{ type: 'paragraph', id: 'b', runs: [] }],
							},
						},
					},
				],
			} as DocumentModel,
			spec,
			newId,
		);
		expect(headerRuns(blank, 0, 'default').map((run) => Boolean(run.image))).toEqual([true, false]);
		expect(headerRuns(blank, 0, 'default')[1]!.text).toBe('Head');
		expect(headerRuns(blank, 0, 'first')[0]!.image!.watermark!.text).toBe('DRAFT');
		const changed = withWatermark(blank, { ...spec, text: 'SAMPLE' }, newId);
		expect(headerRuns(changed, 0, 'default').filter((run) => run.image)).toHaveLength(1);
		expect(currentWatermark(changed)!.text).toBe('SAMPLE');
		const removed = withWatermark(changed, undefined, newId);
		expect(currentWatermark(removed)).toBeUndefined();
		expect(headerRuns(removed, 0, 'default').map((run) => run.text)).toEqual(['Head']);
	});
});

describe('watermark layer', () => {
	const content = (watermark: WatermarkSpec) => ({
		blocks: [
			{
				type: 'paragraph' as const,
				id: 'p',
				runs: [
					{
						text: '',
						image: {
							relId: '',
							partName: '',
							contentType: '',
							widthPx: 0,
							heightPx: 0,
							unsupported: 'Watermark',
							watermark,
						},
					},
				],
			},
		],
	});
	it('draws a centered rotated, half-transparent layer behind the text', () => {
		const layer = watermarkElement(content(spec), { widthPx: 816, heightPx: 1056 })!;
		expect(layer.textContent).toBe('DRAFT');
		expect(layer.style.transform).toBe('rotate(-45deg)');
		expect(layer.style.opacity).toBe('0.5');
		expect(layer.style.pointerEvents).toBe('none');
		expect(layer.style.zIndex).toBe('0');
		const horizontal = watermarkElement(
			content({ ...spec, layout: 'horizontal', semitransparent: false }),
			{
				widthPx: 816,
				heightPx: 1056,
			},
		)!;
		expect(horizontal.style.transform).toBe('none');
		expect(horizontal.style.opacity).toBe('1');
		expect(watermarkElement({ blocks: [] }, { widthPx: 816, heightPx: 1056 })).toBeNull();
	});
});

describe('Watermark dialog', () => {
	it('applies the text and options, and None removes it', () => {
		const applied: Array<WatermarkSpec | undefined> = [];
		const dialog = createWatermarkDialog({
			current: () => undefined,
			canEdit: () => true,
			apply: (value) => applied.push(value),
			restoreFocus: () => {},
		});
		document.body.append(dialog.element);
		dialog.open();
		const field = (name: string) =>
			dialog.element.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${name}"]`)!;
		const ok = [...dialog.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!;
		expect(field('Setting').value).toBe('none');
		field('Setting').value = 'text';
		field('Setting').dispatchEvent(new Event('change'));
		expect(ok.disabled).toBe(true);
		field('Text').value = ' CONFIDENTIAL ';
		field('Text').dispatchEvent(new Event('input'));
		field('Layout').value = 'horizontal';
		field('Semitransparent').checked = false;
		ok.click();
		expect(applied[0]).toMatchObject({
			text: 'CONFIDENTIAL',
			layout: 'horizontal',
			semitransparent: false,
			fontFamily: 'Calibri',
		});
		const again = createWatermarkDialog({
			current: () => applied[0],
			canEdit: () => true,
			apply: (value) => applied.push(value),
			restoreFocus: () => {},
		});
		document.body.append(again.element);
		again.open();
		const second = again.element.querySelector<HTMLSelectElement>('[aria-label="Setting"]')!;
		expect(second.value).toBe('text');
		second.value = 'none';
		second.dispatchEvent(new Event('change'));
		[...again.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!.click();
		expect(applied[1]).toBeUndefined();
	});
});
