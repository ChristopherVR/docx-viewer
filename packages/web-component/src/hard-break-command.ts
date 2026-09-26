import type { Command } from 'prosemirror-state';
import { schema } from './schema';

/** Replaces the current selection with an inline Word-style line break. */
export const insertHardBreak: Command = (state, dispatch, view) => {
	if (view && !view.editable) return false;
	if (!dispatch) return true;
	const transaction = state.tr
		.replaceSelectionWith(schema.nodes.hardBreak.create(), true)
		.scrollIntoView();
	dispatch(transaction);
	return true;
};
