// @vitest-environment jsdom
import { createDocument } from 'docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { shortcutHint } from './binding-labels';
import { editorBindings } from './editor-commands';
import { paragraphStylesPlugin } from './paragraph-styles';
import { syncParagraphToggles } from './paragraph-toggle-sync';
import { createRibbon, setRibbonLocale } from './ribbon';
import { schema } from './schema';

function editor(...paragraphs: Array<Record<string, unknown>>) {
	const doc = schema.node(
		'doc',
		null,
		paragraphs.map((attrs, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}`, ...attrs }, schema.text(`text ${index}`)),
		),
	);
	const model = createDocument();
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema, plugins: [paragraphStylesPlugin(() => model)] }),
	});
	return { view, model };
}
const pressed = (ribbon: HTMLElement, label: string) =>
	ribbon.querySelector(`[aria-label="${label}"]`)!.getAttribute('aria-pressed');
const selectAll = (view: EditorView) =>
	view.dispatch(
		view.state.tr.setSelection(
			TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1),
		),
	);

describe('alignment and list toggles', () => {
	it('presses the alignment of the caret paragraph', () => {
		const ribbon = createRibbon();
		const { view, model } = editor({ align: 'center' });
		syncParagraphToggles(ribbon, view, model);
		expect(pressed(ribbon, 'Align center')).toBe('true');
		expect(pressed(ribbon, 'Align left')).toBe('false');
		expect(pressed(ribbon, 'Justify')).toBe('false');
	});

	it('treats an unaligned paragraph as left and leaves all off when paragraphs differ', () => {
		const ribbon = createRibbon();
		const { view, model } = editor({}, { align: 'right' });
		syncParagraphToggles(ribbon, view, model);
		expect(pressed(ribbon, 'Align left')).toBe('true');
		selectAll(view);
		syncParagraphToggles(ribbon, view, model);
		for (const label of ['Align left', 'Align center', 'Align right', 'Justify'])
			expect(pressed(ribbon, label)).toBe('false');
	});

	it('presses the list buttons only for a list paragraph', () => {
		const ribbon = createRibbon();
		const { view, model } = editor({ numId: 7 });
		syncParagraphToggles(ribbon, view, model);
		expect(pressed(ribbon, 'Bulleted list')).toBe('false');
		const plain = editor({});
		syncParagraphToggles(ribbon, plain.view, plain.model);
		expect(pressed(ribbon, 'Numbered list')).toBe('false');
	});

	it('disables Cut and Copy without a selection and enables them with one', () => {
		const ribbon = createRibbon();
		const { view, model } = editor({});
		syncParagraphToggles(ribbon, view, model);
		expect(ribbon.querySelector<HTMLButtonElement>('[aria-label="Cut"]')!.disabled).toBe(true);
		expect(ribbon.querySelector<HTMLButtonElement>('[aria-label="Copy"]')!.disabled).toBe(true);
		selectAll(view);
		for (const button of ribbon.querySelectorAll<HTMLButtonElement>(
			'[aria-label="Cut"],[aria-label="Copy"]',
		))
			button.disabled = false;
		syncParagraphToggles(ribbon, view, model);
		expect(ribbon.querySelector<HTMLButtonElement>('[aria-label="Copy"]')!.disabled).toBe(false);
	});
});

describe('shortcuts', () => {
	it("binds Word's formatting shortcuts", () => {
		for (const key of [
			'Mod-l',
			'Mod-e',
			'Mod-r',
			'Mod-j',
			'Mod-Shift-.',
			'Mod-Shift-,',
			'Mod-=',
			'Mod-Shift-=',
			'Mod-Space',
		])
			expect(editorBindings[key], key).toBeTypeOf('function');
	});

	it('applies alignment from the keyboard', () => {
		const { view } = editor({});
		editorBindings['Mod-e']!(view.state, view.dispatch, view);
		expect(view.state.doc.firstChild!.attrs.align).toBe('center');
	});

	it('appends the shortcut to ribbon tooltips, and localizes the name', () => {
		const ribbon = createRibbon();
		expect(ribbon.querySelector('[aria-label="Align center"]')!.getAttribute('title')).toBe(
			'Align center (Ctrl+E)',
		);
		expect(ribbon.querySelector('[aria-label="Find and replace"]')!.getAttribute('title')).toBe(
			'Find and replace (Ctrl+F)',
		);
		setRibbonLocale(ribbon, 'fr');
		expect(ribbon.querySelector('[aria-label="Gras"]')!.getAttribute('title')).toBe(
			'Gras (Ctrl+B)',
		);
		expect(shortcutHint('Bold', true)).toBe('⌘B');
		expect(shortcutHint('Insert table')).toBeUndefined();
	});
});
