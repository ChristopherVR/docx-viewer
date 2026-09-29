// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { indentInches, setIndent, syncIndentInputs } from './indent-commands';
import { createRibbon } from './ribbon';
import { schema } from './schema';

function editor(attrs: Array<Record<string, unknown>>) {
	const doc = schema.node(
		'doc',
		null,
		attrs.map((extra, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}`, ...extra }, schema.text(`para ${index}`)),
		),
	);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
}
const selectAll = (view: EditorView) =>
	view.dispatch(
		view.state.tr.setSelection(
			TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1),
		),
	);

describe('indent commands', () => {
	it('sets left and right indents in twips on every selected paragraph', () => {
		const view = editor([{}, {}]);
		selectAll(view);
		expect(setIndent(view, 'left', 0.5)).toBe(true);
		expect(setIndent(view, 'right', 1)).toBe(true);
		view.state.doc.forEach((p) => {
			expect(p.attrs.indentLeftTwips).toBe(720);
			expect(p.attrs.indentRightTwips).toBe(1440);
		});
	});

	it('writes the logical attribute when the paragraph already uses it', () => {
		const view = editor([{ indentStartTwips: 360 }]);
		setIndent(view, 'left', 1);
		expect(view.state.doc.firstChild!.attrs.indentStartTwips).toBe(1440);
		expect(view.state.doc.firstChild!.attrs.indentLeftTwips).toBeNull();
	});

	it("clamps to Word's range, refuses bad numbers and read-only views", () => {
		const view = editor([{}]);
		setIndent(view, 'left', -3);
		expect(view.state.doc.firstChild!.attrs.indentLeftTwips).toBe(0);
		setIndent(view, 'left', 999);
		expect(view.state.doc.firstChild!.attrs.indentLeftTwips).toBe(22 * 1440);
		expect(setIndent(view, 'left', Number.NaN)).toBe(false);
		const locked = new EditorView(document.createElement('div'), {
			state: view.state,
			editable: () => false,
		});
		expect(setIndent(locked, 'left', 1)).toBe(false);
	});

	it('reports a shared indent, or null when paragraphs differ', () => {
		const view = editor([
			{ indentLeftTwips: 720 },
			{ indentLeftTwips: 720 },
			{ indentLeftTwips: 0 },
		]);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 14)));
		expect(indentInches(view.state, 'left')).toBe(0.5);
		selectAll(view);
		expect(indentInches(view.state, 'left')).toBeNull();
	});

	it('shows the indent in the ribbon spinner and emits a change action', () => {
		const ribbon = createRibbon();
		const view = editor([{ indentLeftTwips: 1080 }]);
		syncIndentInputs(ribbon, view.state);
		const left = ribbon.querySelector<HTMLInputElement>('input[aria-label="Indent left"]')!;
		expect(left.value).toBe('0.75');
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		left.value = '1.5';
		left.dispatchEvent(new Event('change', { bubbles: true }));
		expect(seen).toEqual([{ type: 'indent', side: 'left', inches: 1.5 }]);
	});
});
