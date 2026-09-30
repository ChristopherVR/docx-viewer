import type { EditorView } from 'prosemirror-view';
import type { Transaction } from 'prosemirror-state';

/** Runs Font's existing commands against a draft state, then dispatches one atomic change. */
export function batchFontFormat(view: EditorView, apply: (draft: EditorView) => void): void {
	let state = view.state;
	const combined = state.tr;
	let scroll = false;
	const dispatch = (transaction: Transaction) => {
		for (const step of transaction.steps) combined.step(step);
		state = state.apply(transaction);
		scroll ||= transaction.scrolledIntoView;
	};
	const draft = new Proxy(view, {
		get(target, key) {
			if (key === 'state') return state;
			if (key === 'dispatch') return dispatch;
			return Reflect.get(target, key, target);
		},
	});
	apply(draft);
	combined
		.setSelection(state.selection.getBookmark().resolve(combined.doc))
		.setStoredMarks(state.storedMarks);
	if (scroll) combined.scrollIntoView();
	if (combined.steps.length || combined.storedMarksSet) view.dispatch(combined);
}
