// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { applyHighlight, toggleVerticalAlign } from './inline-commands';
import { updateParagraphs } from './ribbon-commands';

let view: EditorView;
afterEach(() => {
	view?.destroy();
	document.body.replaceChildren();
});
function editor() {
	const doc = schema.nodes.doc.create(
		null,
		schema.nodes.paragraph.create({ id: 'p', indentLeftTwips: 360 }, schema.text('Testing')),
	);
	view = new EditorView(document.body, {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 1, 5) }),
	});
	return view;
}
it('switches superscript to subscript without losing highlights or unselected text', () => {
	editor();
	applyHighlight(view, 'yellow');
	toggleVerticalAlign(view, 'superscript');
	toggleVerticalAlign(view, 'subscript');
	const first = view.state.doc.firstChild!.firstChild!;
	expect(first.text).toBe('Test');
	expect(first.marks.map((mark) => [mark.type.name, mark.attrs])).toEqual([
		['highlight', { color: 'yellow' }],
		['verticalAlign', { value: 'subscript' }],
	]);
	expect(view.state.doc.firstChild!.lastChild!.marks).toHaveLength(0);
	toggleVerticalAlign(view, 'subscript');
	expect(view.state.doc.firstChild!.firstChild!.marks.map((mark) => mark.type.name)).toEqual([
		'highlight',
	]);
});
it('retains cursor formatting until typing and refuses invalid/read-only changes', () => {
	editor();
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
	applyHighlight(view, 'cyan');
	toggleVerticalAlign(view, 'superscript');
	view.dispatch(view.state.tr.insertText('New'));
	expect(view.state.doc.firstChild!.firstChild!.marks).toHaveLength(2);
	const before = view.state.doc.toJSON();
	applyHighlight(view, 'invalid-css; color:red');
	view.setProps({ editable: () => false });
	toggleVerticalAlign(view, 'subscript');
	expect(view.state.doc.toJSON()).toEqual(before);
});
it('restores inherited spacing without NaN and retains explicit zero indentation', () => {
	editor();
	updateParagraphs(view, 'spacingAfter', '240');
	updateParagraphs(view, 'spacingAfter', 'inherit');
	updateParagraphs(view, 'indent', 'decrease');
	expect(view.state.doc.firstChild!.attrs.spacingAfterTwips).toBeNull();
	expect(view.state.doc.firstChild!.attrs.indentLeftTwips).toBe(0);
	updateParagraphs(view, 'spacingBefore', 'mixed');
	expect(view.state.doc.firstChild!.attrs.spacingBeforeTwips).toBeNull();
});
