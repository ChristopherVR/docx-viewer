// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { readDropCap, setDropCap } from './drop-cap-command';
import { history, undo, redo } from 'prosemirror-history';
import { paragraphStyle } from './schema';
import { schema } from './schema';

function editor(text: string) {
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create({ id: 'p' }, schema.text(text)),
	]);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema, plugins: [history()] }),
	});
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 3)));
	return view;
}

describe('setDropCap', () => {
	it('applies font, line count and distance together, preserves them on a preset change, and undoes', () => {
		const view = editor('Once upon a time');
		expect(setDropCap(view, 'drop', { lines: 4, distanceTwips: 144, fontFamily: 'Georgia' })).toBe(
			true,
		);
		expect(readDropCap(view)).toEqual({
			style: 'drop',
			lines: 4,
			distanceTwips: 144,
			fontFamily: 'Georgia',
		});
		expect(paragraphStyle(view.state.doc.firstChild!.attrs)).toContain('margin-right:9.6px');
		setDropCap(view, 'margin');
		expect(readDropCap(view)).toMatchObject({ style: 'margin', lines: 4, distanceTwips: 144 });
		setDropCap(view, 'drop', { lines: 2, distanceTwips: 0, fontFamily: 'Arial' });
		expect(readDropCap(view)).toEqual({
			style: 'drop',
			lines: 2,
			distanceTwips: 0,
			fontFamily: 'Arial',
		});
		undo(view.state, view.dispatch);
		expect(readDropCap(view)).toMatchObject({ style: 'margin', lines: 4, fontFamily: 'Georgia' });
		redo(view.state, view.dispatch);
		expect(readDropCap(view)).toMatchObject({ lines: 2, fontFamily: 'Arial' });
		expect(setDropCap(view, 'drop', { lines: 1.5 })).toBe(false);
		view.destroy();
	});
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
