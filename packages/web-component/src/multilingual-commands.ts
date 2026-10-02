import type { Mark } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { isValidLanguageTag } from 'docx-core';
import { schema } from './schema';

export type LanguageField = 'language' | 'eastAsiaLanguage' | 'bidiLanguage';
export type MultilingualAction =
	| { type: 'paragraphDirection'; value: 'inherit' | 'ltr' | 'rtl' }
	| { type: 'language'; key: LanguageField; value: string }
	| { type: 'runRtl'; value: 'inherit' | 'on' | 'off' };

function languageMarkAttrs(mark: Mark | undefined, key: LanguageField, value: string) {
	const attrs = {
		language: mark?.attrs.language ?? null,
		eastAsiaLanguage: mark?.attrs.eastAsiaLanguage ?? null,
		bidiLanguage: mark?.attrs.bidiLanguage ?? null,
	};
	attrs[key] = value || null;
	return attrs;
}

function applyLanguage(view: EditorView, key: LanguageField, value: string): void {
	const tag = value.trim();
	if (tag && !isValidLanguageTag(tag)) return;
	const { state } = view;
	const type = schema.marks.language;
	if (state.selection.empty) {
		const marks = state.storedMarks || state.selection.$from.marks();
		const existing = marks.find((mark) => mark.type === type);
		const attrs = languageMarkAttrs(existing, key, tag);
		const next = Object.values(attrs).some(Boolean) ? type.create(attrs) : undefined;
		view.dispatch(
			state.tr.setStoredMarks([
				...marks.filter((mark) => mark.type !== type),
				...(next ? [next] : []),
			]),
		);
		return;
	}
	let transaction = state.tr;
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (!node.isText) return;
		const start = Math.max(pos, state.selection.from);
		const end = Math.min(pos + node.nodeSize, state.selection.to);
		const existing = node.marks.find((mark) => mark.type === type);
		const attrs = languageMarkAttrs(existing, key, tag);
		const next = Object.values(attrs).some(Boolean) ? type.create(attrs) : undefined;
		transaction = transaction.removeMark(start, end, type);
		if (next) transaction = transaction.addMark(start, end, next);
	});
	view.dispatch(closeHistory(transaction));
}

function applyRunRtl(view: EditorView, value: 'inherit' | 'on' | 'off'): void {
	const { state } = view;
	const type = schema.marks.runRtl;
	const mark = value === 'inherit' ? undefined : type.create({ value: value === 'on' });
	if (state.selection.empty) {
		const marks = state.storedMarks || state.selection.$from.marks();
		view.dispatch(
			state.tr.setStoredMarks([
				...marks.filter((item) => item.type !== type),
				...(mark ? [mark] : []),
			]),
		);
		return;
	}
	const transaction = state.tr.removeMark(state.selection.from, state.selection.to, type);
	view.dispatch(
		closeHistory(
			mark ? transaction.addMark(state.selection.from, state.selection.to, mark) : transaction,
		),
	);
}

function applyDirection(view: EditorView, value: 'inherit' | 'ltr' | 'rtl'): void {
	const { state } = view;
	const positions: number[] = [];
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (node.type.name === 'paragraph') positions.push(pos);
	});
	if (!positions.length && state.selection.$from.parent.type.name === 'paragraph')
		positions.push(state.selection.$from.before());
	let transaction = state.tr;
	for (const pos of positions) {
		const node = transaction.doc.nodeAt(pos);
		if (node)
			transaction = transaction.setNodeMarkup(pos, undefined, {
				...node.attrs,
				direction: value === 'inherit' ? null : value,
			});
	}
	if (transaction.docChanged) view.dispatch(closeHistory(transaction).scrollIntoView());
}

export function applyMultilingualAction(view: EditorView, action: MultilingualAction): void {
	if (!view.editable) return;
	if (action.type === 'language') applyLanguage(view, action.key, action.value);
	else if (action.type === 'runRtl') applyRunRtl(view, action.value);
	else applyDirection(view, action.value);
}
