// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import { EditorState } from 'prosemirror-state';
import { docToModel, modelToDoc } from './model-adapter';
import { createRibbon, type RibbonAction } from './ribbon';
import { syncMultilingualControls } from './multilingual-ribbon';
import { schema } from './schema';

describe('multilingual editor metadata', () => {
	it('roundtrips script language tags, direct run direction, paragraph direction and Unicode', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'rtl',
				direction: 'rtl',
				runs: [
					{
						text: 'مرحبا עברית 中文 e\u0301 😀',
						language: 'ar-SA',
						eastAsiaLanguage: 'ja-JP',
						bidiLanguage: 'ar-SA',
						rtl: true,
					},
					{ text: 'explicit off', rtl: false },
				],
			},
		];
		const editorDoc = modelToDoc(model);
		const updated = docToModel(editorDoc, model);
		expect(updated.blocks).toEqual(model.blocks);
		const paragraphDom = schema.nodes.paragraph.spec.toDOM!(
			editorDoc.firstChild!,
		) as unknown as readonly unknown[];
		expect(paragraphDom[1]).toMatchObject({ dir: 'rtl' });
		expect(String((paragraphDom[1] as { style: string }).style)).toContain('direction:rtl');
		expect(String((paragraphDom[1] as { style: string }).style)).not.toContain('text-align:left');
	});

	it('exposes accessible language and direction controls with BCP 47 presets', () => {
		const ribbon = createRibbon();
		ribbon
			.querySelector<HTMLButtonElement>('[role="tab"][aria-controls="dve-panel-review"]')!
			.click();
		const labels = [
			'Paragraph direction',
			'Text language',
			'East Asian language',
			'Complex script language',
			'Run direction',
		];
		for (const label of labels)
			expect(ribbon.querySelector(`[aria-label="${label}"]`)).toBeTruthy();
		expect(
			[...ribbon.querySelector<HTMLSelectElement>('[aria-label="Text language"]')!.options].some(
				(option) => option.value === 'ar-SA',
			),
		).toBe(true);
		expect(ribbon.querySelector('[aria-label="Find and replace"]')).toBeTruthy();
	});

	it('syncs direction, language, and explicit run RTL values from editor selection', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'rtl',
				direction: 'rtl',
				runs: [{ text: 'Arabic', language: 'ar-SA', bidiLanguage: 'ar-SA', rtl: false }],
			},
		];
		const state = EditorState.create({ doc: modelToDoc(model) });
		const ribbon = createRibbon();
		ribbon
			.querySelector<HTMLButtonElement>('[role="tab"][aria-controls="dve-panel-review"]')!
			.click();
		syncMultilingualControls(ribbon, state);
		expect(
			ribbon.querySelector<HTMLSelectElement>('[aria-label="Paragraph direction"]')!.value,
		).toBe('rtl');
		expect(ribbon.querySelector<HTMLSelectElement>('[aria-label="Text language"]')!.value).toBe(
			'ar-SA',
		);
		expect(ribbon.querySelector<HTMLSelectElement>('[aria-label="Run direction"]')!.value).toBe(
			'off',
		);
	});

	it('dispatches typed multilingual and search actions', () => {
		const ribbon = createRibbon();
		ribbon
			.querySelector<HTMLButtonElement>('[role="tab"][aria-controls="dve-panel-review"]')!
			.click();
		const actions: RibbonAction[] = [];
		ribbon.addEventListener('ribbon-action', (event) =>
			actions.push((event as CustomEvent<RibbonAction>).detail),
		);
		const language = ribbon.querySelector<HTMLSelectElement>('[aria-label="Text language"]')!;
		language.value = 'ar-SA';
		language.dispatchEvent(new Event('change', { bubbles: true }));
		const direction = ribbon.querySelector<HTMLSelectElement>(
			'[aria-label="Paragraph direction"]',
		)!;
		direction.value = 'rtl';
		direction.dispatchEvent(new Event('change', { bubbles: true }));
		ribbon.querySelector<HTMLButtonElement>('[aria-label="Find and replace"]')!.click();
		expect(actions).toContainEqual({ type: 'language', key: 'language', value: 'ar-SA' });
		expect(actions).toContainEqual({ type: 'paragraphDirection', value: 'rtl' });
		expect(actions).toContainEqual({ type: 'search' });
	});
});
