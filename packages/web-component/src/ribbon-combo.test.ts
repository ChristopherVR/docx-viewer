// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { afterEach, describe, expect, it } from 'vitest';
import { syncFontControls } from './font-sync';
import { runStylesPlugin } from './run-styles';
import { parseFontFamily, parseFontSize, setComboValue, type ComboInput } from './ribbon-combo';
import { createRibbon } from './ribbon';
import { schema } from './schema';

afterEach(() => (document.body.innerHTML = ''));

const combo = (ribbon: HTMLElement, label: string) =>
	ribbon.querySelector<ComboInput>(`input[aria-label="${label}"]`)!;

describe('font input parsing', () => {
	it('accepts point sizes from 1 to 1638 in half points and comma decimals', () => {
		expect(parseFontSize('12')).toBe(12);
		expect(parseFontSize(' 10,5 ')).toBe(10.5);
		expect(parseFontSize('10.26')).toBe(10.5);
		expect(parseFontSize('0.4')).toBeNull();
		expect(parseFontSize('1639')).toBeNull();
		expect(parseFontSize('big')).toBeNull();
		expect(parseFontSize('')).toBeNull();
	});

	it('accepts a single plain family name and rejects markup or CSS punctuation', () => {
		expect(parseFontFamily('  Times   New Roman ')).toBe('Times New Roman');
		expect(parseFontFamily('')).toBeNull();
		expect(parseFontFamily('Arial; color:red')).toBeNull();
		expect(parseFontFamily('"Arial"')).toBeNull();
		expect(parseFontFamily('x'.repeat(65))).toBeNull();
	});
});

describe('font combo boxes', () => {
	it('apply a typed value on Enter and emit the ribbon action', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		const size = combo(ribbon, 'Font size');
		size.value = '13,5';
		size.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(seen).toEqual([{ type: 'font', key: 'size', value: '13.5' }]);
		expect(size.value).toBe('13.5');
	});

	it('revert an invalid entry instead of applying it', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		const size = combo(ribbon, 'Font size');
		expect(size.value).toBe('11');
		size.value = 'huge';
		size.dispatchEvent(new Event('change', { bubbles: true }));
		expect(seen).toEqual([]);
		expect(size.value).toBe('11');
	});

	it('open a list, mark the current entry and apply a picked font', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const ribbon = createRibbon();
		host.append(ribbon);
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		const wrap = combo(ribbon, 'Font family').parentElement!;
		wrap.querySelector<HTMLButtonElement>('.ribbon-combo-caret')!.click();
		const options = [
			...document.querySelectorAll<HTMLButtonElement>('.ribbon-list [role="option"]'),
		];
		expect(options.length).toBeGreaterThan(15);
		expect(
			options.find((option) => option.getAttribute('aria-selected') === 'true')?.textContent,
		).toBe('Calibri');
		options.find((option) => option.textContent === 'Georgia')!.click();
		expect(seen).toEqual([{ type: 'font', key: 'family', value: 'Georgia' }]);
		expect(document.querySelector('.ribbon-list')).toBeNull();
	});

	it('list a value the document uses that is not in the standard list', () => {
		const input = combo(createRibbon(), 'Font family');
		setComboValue(input, 'Garamond Premier Pro');
		expect(input.comboItems).toContain('Garamond Premier Pro');
		expect(input.value).toBe('Garamond Premier Pro');
	});
});

describe('font sync shows the effective font', () => {
	const doc = (
		attrs: Record<string, unknown>,
		...marks: ReturnType<typeof schema.marks.bold.create>[]
	) =>
		schema.node('doc', null, [
			schema.nodes.paragraph!.create({ id: 'p', ...attrs }, schema.text('some text', marks)),
		]);

	it('reads direct formatting, blank when a selection mixes fonts', () => {
		const ribbon = createRibbon();
		const model = createDocument();
		const font = (family: string, size: number) => schema.marks.font!.create({ family, size });
		const mixed = schema.node('doc', null, [
			schema.nodes.paragraph!.create({ id: 'p' }, [
				schema.text('one ', [font('Georgia', 14)]),
				schema.text('two', [font('Verdana', 14)]),
			]),
		]);
		let state = EditorState.create({ doc: mixed, schema, plugins: [runStylesPlugin(() => model)] });
		state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 3)));
		syncFontControls(ribbon, state);
		expect(combo(ribbon, 'Font family').value).toBe('Georgia');
		expect(combo(ribbon, 'Font size').value).toBe('14');
		state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, 8)));
		syncFontControls(ribbon, state);
		expect(combo(ribbon, 'Font family').value).toBe('');
		expect(combo(ribbon, 'Font size').value).toBe('14');
	});

	it('falls back to the paragraph style font when the run has none', () => {
		const ribbon = createRibbon();
		const model = createDocument();
		model.characterStyles = {
			docDefaults: {},
			warnings: [],
			styles: {
				Big: { id: 'Big', type: 'paragraph', formatting: { fontFamily: 'Cambria', fontSize: 20 } },
			},
		};
		model.paragraphStyles = {
			docDefaults: {},
			warnings: [],
			styles: { Big: { id: 'Big', formatting: {} } },
		};
		let state = EditorState.create({
			doc: doc({ style: 'Big' }),
			schema,
			plugins: [runStylesPlugin(() => model)],
		});
		state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 2)));
		syncFontControls(ribbon, state);
		expect(combo(ribbon, 'Font family').value).toBe('Cambria');
		expect(combo(ribbon, 'Font size').value).toBe('20');
	});
});
