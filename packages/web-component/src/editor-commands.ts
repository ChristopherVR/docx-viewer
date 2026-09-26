import type { Command } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { undo, redo } from 'prosemirror-history';
import { baseKeymap, toggleMark } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import type { RibbonAction } from './ribbon';
import { schema } from './schema';
import {
	applyFont,
	clearFormatting,
	insertTable,
	updateParagraphs,
	updatePage,
} from './ribbon-commands';
import { applyHighlight, toggleVerticalAlign } from './inline-commands';
import { executeTableCommand } from './table-commands';
import { insertHardBreak } from './hard-break-command';
import { applyMultilingualAction } from './multilingual-ribbon';

const marks = {
	bold: toggleMark(schema.marks.bold),
	italic: toggleMark(schema.marks.italic),
	underline: toggleMark(schema.marks.underline),
	strike: toggleMark(schema.marks.strike),
};
const editable =
	(command: Command): Command =>
	(state, dispatch, view) =>
		view?.editable === false ? false : command(state, dispatch, view);

export function editorKeymap(showSearch: () => void) {
	const shortcuts = Object.fromEntries(
		Object.entries({
			...baseKeymap,
			'Mod-z': undo,
			'Mod-y': redo,
			'Mod-Shift-z': redo,
			'Mod-b': marks.bold,
			'Mod-i': marks.italic,
			'Mod-u': marks.underline,
			'Shift-Enter': insertHardBreak,
		}).map(([key, command]) => [key, editable(command)]),
	);
	return keymap({
		...shortcuts,
		'Mod-f': () => {
			showSearch();
			return true;
		},
	});
}

export function runRibbonCommand(
	view: EditorView,
	action: RibbonAction,
	nextId?: (kind: string) => string,
) {
	if (!view.editable) return;
	if (
		action.type === 'paragraphDirection' ||
		action.type === 'language' ||
		action.type === 'runRtl'
	)
		applyMultilingualAction(view, action);
	else if (action.type === 'format') {
		if (action.key === 'superscript' || action.key === 'subscript')
			toggleVerticalAlign(view, action.key);
		else marks[action.key](view.state, view.dispatch, view);
	} else if (action.type === 'align') {
		let transaction = view.state.tr;
		view.state.doc.nodesBetween(view.state.selection.from, view.state.selection.to, (node, pos) => {
			if (node.type.name === 'paragraph')
				transaction = transaction.setNodeMarkup(pos, undefined, {
					...node.attrs,
					align: action.value,
				});
		});
		if (transaction.docChanged) view.dispatch(transaction);
	} else if (action.type === 'history')
		(action.key === 'undo' ? undo : redo)(view.state, view.dispatch, view);
	else if (action.type === 'font') {
		if (action.key === 'highlight') applyHighlight(view, action.value);
		else applyFont(view, action.key, action.value);
	} else if (action.type === 'tableEdit' && !nextId) executeTableCommand(view, action.key);
	else if (action.type === 'clear') clearFormatting(view);
	else if (action.type === 'table') insertTable(view, nextId);
	else if (action.type === 'page') updatePage(view, action.key, action.value);
	else if (action.type === 'paragraph') updateParagraphs(view, action.key, action.value);
}
