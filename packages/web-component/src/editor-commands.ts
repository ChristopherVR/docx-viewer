import type { Command } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { undo, redo } from 'prosemirror-history';
import { baseKeymap, chainCommands } from 'prosemirror-commands';
import { expectDefined } from './defined';
import { toggleFormat } from './toggle-commands';
import { keymap } from 'prosemirror-keymap';
import type { RibbonAction } from './ribbon';
import { applyFont, clearFormatting, insertTable, updateParagraphs } from './ribbon-commands';
import { applyHighlight, toggleVerticalAlign } from './inline-commands';
import { executeTableCommand } from './table-commands';
import { insertHardBreak } from './hard-break-command';
import { insertPageBreak, insertBreakCommand } from './page-break-command';
import { applyMultilingualAction } from './multilingual-ribbon';
import { stepFontSize } from './font-step';
import { changeCase } from './change-case';
import { setIndent } from './indent-commands';
import { sortParagraphs } from './sort-commands';
import { setBorders, setShading } from './paragraph-decoration';
import { formatDateTime, insertPlainText } from './insert-text-commands';
import { selectAll } from 'prosemirror-commands';
import { exitListOnEmptyEnter, indentListItem, outdentListItem } from './list-commands';

const marks = {
	bold: toggleFormat('bold'),
	italic: toggleFormat('italic'),
	underline: toggleFormat('underline'),
	strike: toggleFormat('strike'),
};
const editable =
	(command: Command): Command =>
	(state, dispatch, view) =>
		view?.editable === false ? false : command(state, dispatch, view);

/** Runs `action` on an editable view; returns true so the key never reaches the browser. */
function runIfEditable(view: EditorView | undefined, action: (view: EditorView) => void): boolean {
	if (view?.editable) action(view);
	return true;
}

/** Word's Ctrl+L / E / R / J paragraph alignment shortcuts. */
function alignBindings(): Record<string, Command> {
	const bind =
		(value: 'left' | 'center' | 'right' | 'justify'): Command =>
		(_state, _dispatch, view) =>
			runIfEditable(view, (v) => runRibbonCommand(v, { type: 'align', value }));
	return {
		'Mod-l': bind('left'),
		'Mod-e': bind('center'),
		'Mod-r': bind('right'),
		'Mod-j': bind('justify'),
	};
}

/** ProseMirror bindings the editor adds on top of the base keymap; the help dialog lists these. */
export const editorBindings: Record<string, Command> = {
	'Mod-z': undo,
	'Mod-y': redo,
	'Mod-Shift-z': redo,
	'Mod-b': marks.bold,
	'Mod-i': marks.italic,
	'Mod-u': marks.underline,
	'Shift-Enter': insertHardBreak,
	Enter: chainCommands(
		exitListOnEmptyEnter,
		expectDefined(baseKeymap.Enter, 'the base Enter command'),
	),
	Tab: indentListItem,
	'Shift-Tab': outdentListItem,
	'Mod-Enter': insertPageBreak,
	...alignBindings(),
	'Mod-Shift-.': (_state, _dispatch, view) => runIfEditable(view, (v) => stepFontSize(v, 'grow')),
	'Mod-Shift-,': (_state, _dispatch, view) => runIfEditable(view, (v) => stepFontSize(v, 'shrink')),
	'Mod-=': (_state, _dispatch, view) =>
		runIfEditable(view, (v) => toggleVerticalAlign(v, 'subscript')),
	'Mod-Shift-=': (_state, _dispatch, view) =>
		runIfEditable(view, (v) => toggleVerticalAlign(v, 'superscript')),
	'Mod-Space': (_state, _dispatch, view) => runIfEditable(view, clearFormatting),
};

export function editorKeymap(showSearch: () => void) {
	const shortcuts = Object.fromEntries(
		Object.entries({ ...baseKeymap, ...editorBindings }).map(([key, command]) => [
			key,
			editable(command),
		]),
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
	else if (action.type === 'sort')
		sortParagraphs(
			view,
			action.order,
			view.dom.closest('[lang]')?.getAttribute('lang') ?? undefined,
		);
	else if (action.type === 'blankPage') {
		// Word's Blank Page is two page breaks: one ends this page, one ends the empty page.
		insertPageBreak(view.state, view.dispatch, view);
		insertPageBreak(view.state, view.dispatch, view);
	} else if (action.type === 'shading')
		setShading(view, action.value === 'none' ? null : action.value);
	else if (action.type === 'borders') setBorders(view, action.preset);
	else if (action.type === 'indent') setIndent(view, action.side, action.inches);
	else if (action.type === 'insertSymbol') insertPlainText(view, action.value);
	else if (action.type === 'insertDateTime')
		insertPlainText(
			view,
			formatDateTime(
				action.value,
				view.dom.closest('[lang]')?.getAttribute('lang') ?? navigator.language,
			),
		);
	else if (action.type === 'changeCase') changeCase(view, action.value);
	else if (action.type === 'fontStep') stepFontSize(view, action.direction);
	else if (action.type === 'clear') clearFormatting(view);
	else if (action.type === 'table') insertTable(view, nextId);
	else if (action.type === 'insertBreak')
		insertBreakCommand(action.kind)(view.state, view.dispatch, view);
	else if (action.type === 'paragraph') updateParagraphs(view, action.key, action.value);
}
