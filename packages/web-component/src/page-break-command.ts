import type { Command } from 'prosemirror-state';
import { schema } from './schema';

/** Inserts a visible, editable Word page or column break marker (Ctrl+Enter for a page break). */
export function insertBreakCommand(kind: 'page' | 'column'): Command {
	return (state, dispatch, view) => {
		if (view && !view.editable) return false;
		if (!dispatch) return true;
		const transaction = state.tr
			.replaceSelectionWith(schema.nodes.pageBreak.create({ kind }), false)
			.scrollIntoView();
		dispatch(transaction);
		return true;
	};
}

export const insertPageBreak: Command = insertBreakCommand('page');
export const insertColumnBreak: Command = insertBreakCommand('column');
