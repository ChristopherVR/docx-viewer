import type { MarkType } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { schema, wordHighlightColors } from './schema';
import { selectionScript } from './script-state';

function setInlineMark(view: EditorView, type: MarkType, attrs?: Record<string, string>) {
	if (!view.editable) return;
	const { state } = view;
	const { from, to, empty } = state.selection;
	const mark = attrs ? type.create(attrs) : undefined;
	if (empty) {
		const marks = (state.storedMarks ?? state.selection.$from.marks()).filter(
			(item) => item.type !== type,
		);
		view.dispatch(state.tr.setStoredMarks(mark ? [...marks, mark] : marks));
	} else {
		const transaction = state.tr.removeMark(from, to, type);
		view.dispatch(mark ? transaction.addMark(from, to, mark) : transaction);
	}
}

export function applyHighlight(view: EditorView, color: string) {
	if (color !== 'none' && !Object.hasOwn(wordHighlightColors, color)) return;
	setInlineMark(view, schema.marks.highlight, color === 'none' ? undefined : { color });
}

export function toggleVerticalAlign(view: EditorView, value: 'superscript' | 'subscript') {
	applyVerticalAlign(view, selectionScript(view.state) === value ? 'baseline' : value);
}

export function applyVerticalAlign(
	view: EditorView,
	value: 'baseline' | 'superscript' | 'subscript',
) {
	setInlineMark(view, schema.marks.verticalAlign, { value });
}
