import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { schema } from './schema';

export interface RevisionRange {
	id: string;
	kind: 'insert' | 'delete';
	from: number;
	to: number;
	author: string;
}

/** Merges adjacent text/hardBreak nodes sharing the same revision mark id into one range. */
export function collectRevisionRanges(doc: import('prosemirror-model').Node): RevisionRange[] {
	const ranges: RevisionRange[] = [];
	doc.descendants((node, pos) => {
		if (!node.isText && node.type.name !== 'hardBreak') return;
		const mark = node.marks.find(
			(item) => item.type.name === 'insertion' || item.type.name === 'deletion',
		);
		if (!mark) return;
		const kind = mark.type.name === 'insertion' ? 'insert' : 'delete';
		const id = String(mark.attrs.id);
		const last = ranges.at(-1);
		if (last && last.id === id && last.kind === kind && last.to === pos)
			last.to = pos + node.nodeSize;
		else
			ranges.push({
				id,
				kind,
				from: pos,
				to: pos + node.nodeSize,
				author: String(mark.attrs.author),
			});
	});
	return ranges;
}

function revisionNear(view: EditorView): RevisionRange | undefined {
	const ranges = collectRevisionRanges(view.state.doc);
	const { from, to } = view.state.selection;
	return (
		ranges.find((range) => range.from < to && range.to > from) ?? ranges.find((r) => r.from >= to)
	);
}

export function hasAnyChange(view: EditorView | undefined): boolean {
	return Boolean(view && collectRevisionRanges(view.state.doc).length);
}
export function hasChangeAtCursor(view: EditorView | undefined): boolean {
	return Boolean(view && revisionNear(view));
}

function markType(kind: 'insert' | 'delete') {
	return kind === 'insert' ? schema.marks.insertion : schema.marks.deletion;
}

function resolveRange(view: EditorView, range: RevisionRange): { from: number; to: number } {
	// Re-locate by scanning current marks with the same id, in case the doc changed since collection.
	const current = collectRevisionRanges(view.state.doc).find((item) => item.id === range.id);
	return current ?? range;
}

/** Accepting an insertion keeps the text; accepting a deletion removes it. */
export function acceptRevisionRange(view: EditorView, range: RevisionRange): void {
	const { from, to } = resolveRange(view, range);
	const tr =
		range.kind === 'insert'
			? view.state.tr.removeMark(from, to, markType('insert'))
			: view.state.tr.delete(from, to);
	view.dispatch(closeHistory(tr).scrollIntoView());
}
/** Rejecting an insertion removes the text; rejecting a deletion restores it. */
export function rejectRevisionRange(view: EditorView, range: RevisionRange): void {
	const { from, to } = resolveRange(view, range);
	const tr =
		range.kind === 'insert'
			? view.state.tr.delete(from, to)
			: view.state.tr.removeMark(from, to, markType('delete'));
	view.dispatch(closeHistory(tr).scrollIntoView());
}

export function acceptChangeAtCursor(view: EditorView): boolean {
	const range = revisionNear(view);
	if (!range) return false;
	acceptRevisionRange(view, range);
	return true;
}
export function rejectChangeAtCursor(view: EditorView): boolean {
	const range = revisionNear(view);
	if (!range) return false;
	rejectRevisionRange(view, range);
	return true;
}

function applyToAll(view: EditorView, mode: 'accept' | 'reject'): boolean {
	const ranges = collectRevisionRanges(view.state.doc);
	if (!ranges.length) return false;
	let tr = view.state.tr;
	for (const range of ranges) {
		const from = tr.mapping.map(range.from);
		const to = tr.mapping.map(range.to);
		const removeText = (mode === 'accept') === (range.kind === 'delete');
		if (removeText) tr = tr.delete(from, to);
		else tr = tr.removeMark(from, to, markType(range.kind));
	}
	view.dispatch(closeHistory(tr));
	return true;
}
export const acceptAllChanges = (view: EditorView): boolean => applyToAll(view, 'accept');
export const rejectAllChanges = (view: EditorView): boolean => applyToAll(view, 'reject');

function navigate(view: EditorView, direction: 'next' | 'previous'): boolean {
	const ranges = collectRevisionRanges(view.state.doc);
	if (!ranges.length) return false;
	const { from, to } = view.state.selection;
	const target =
		direction === 'next'
			? (ranges.find((range) => range.from > to) ?? ranges[0])
			: ([...ranges].reverse().find((range) => range.to < from) ?? ranges.at(-1)!);
	view.dispatch(
		view.state.tr
			.setSelection(TextSelection.create(view.state.doc, target.from, target.to))
			.scrollIntoView(),
	);
	view.focus();
	return true;
}
export const goToNextChange = (view: EditorView): boolean => navigate(view, 'next');
export const goToPreviousChange = (view: EditorView): boolean => navigate(view, 'previous');
