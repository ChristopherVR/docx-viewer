// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { setDropCap } from './drop-cap-command';
import { schema } from './schema';

function editor(text: string) {
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create({ id: 'p' }, schema.text(text)),
	]);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 3)));
	return view;
}

describe('setDropCap', () => {
	it('moves the first letter into an enlarged frame paragraph and folds it back', () => {
		const view = editor('Once upon a time');
		expect(setDropCap(view, 'drop')).toBe(true);
		const [cap, body] = [view.state.doc.child(0), view.state.doc.child(1)];
		expect(cap.textContent).toBe('O');
		expect(cap.attrs.dropCap).toEqual({ style: 'drop', lines: 3 });
		expect(cap.attrs.lineSpacingRule).toBe('exact');
		expect(cap.firstChild!.marks.find((m) => m.type.name === 'font')!.attrs.size).toBe(54);
		expect(body.textContent).toBe('nce upon a time');

		expect(setDropCap(view, 'margin')).toBe(true);
		expect(view.state.doc.child(0).attrs.dropCap.style).toBe('margin');

		expect(setDropCap(view, 'none')).toBe(true);
		expect(view.state.doc.childCount).toBe(1);
		expect(view.state.doc.textContent).toBe('Once upon a time');
		expect(view.state.doc.firstChild!.firstChild!.marks.some((m) => m.type.name === 'font')).toBe(
			false,
		);
	});

	it('does nothing for an empty paragraph or when there is no cap to remove', () => {
		const view = editor('x');
		expect(setDropCap(view, 'none')).toBe(false);
		const empty = new EditorView(document.createElement('div'), {
			state: EditorState.create({
				doc: schema.node('doc', null, [schema.nodes.paragraph!.create()]),
				schema,
			}),
		});
		expect(setDropCap(empty, 'drop')).toBe(false);
	});
});
