import type { Comment } from '@christophervr/docx-core';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { TextSelection } from 'prosemirror-state';
import { schema } from './schema';

let commentSerial = 0;
function nextCommentId(idGenerator?: (kind: string) => string): string {
	return idGenerator
		? idGenerator('comment')
		: `dve-comment-${Date.now().toString(36)}-${++commentSerial}`;
}

/** Wraps the current selection with a `comment` mark carrying a new comment id; returns the new entry. */
export function addComment(
	view: EditorView,
	author: string,
	text: string,
	idGenerator?: (kind: string) => string,
): Comment | null {
	const { from, to } = view.state.selection;
	if (from === to) return null;
	const id = nextCommentId(idGenerator);
	let tr = view.state.tr;
	view.state.doc.nodesBetween(from, to, (node, pos) => {
		if (!node.isText && node.type.name !== 'hardBreak') return;
		const start = Math.max(pos, from);
		const end = Math.min(pos + node.nodeSize, to);
		if (end <= start) return;
		const existing = node.marks.find((mark) => mark.type.name === 'comment');
		const ids = existing ? Array.from(new Set([...(existing.attrs.ids as string[]), id])) : [id];
		if (existing) tr = tr.removeMark(start, end, schema.marks.comment);
		tr = tr.addMark(start, end, schema.marks.comment.create({ ids }));
	});
	view.dispatch(closeHistory(tr));
	return { id, author, text, resolved: false };
}

/** Removes one comment id from every anchor range that carries it. */
export function removeCommentAnchor(view: EditorView, id: string): void {
	let tr = view.state.tr;
	view.state.doc.descendants((node, pos) => {
		if (!node.isText && node.type.name !== 'hardBreak') return;
		const mark = node.marks.find((item) => item.type.name === 'comment');
		if (!mark || !(mark.attrs.ids as string[]).includes(id)) return;
		const remaining = (mark.attrs.ids as string[]).filter((existing) => existing !== id);
		tr = tr.removeMark(pos, pos + node.nodeSize, schema.marks.comment);
		if (remaining.length)
			tr = tr.addMark(pos, pos + node.nodeSize, schema.marks.comment.create({ ids: remaining }));
	});
	if (tr.docChanged) view.dispatch(closeHistory(tr));
}

export function deleteComment(comments: Comment[], id: string): Comment[] {
	return comments.filter((comment) => comment.id !== id && comment.parentId !== id);
}
export function resolveComment(comments: Comment[], id: string, resolved: boolean): Comment[] {
	return comments.map((comment) => (comment.id === id ? { ...comment, resolved } : comment));
}
export function replyToComment(
	comments: Comment[],
	parentId: string,
	author: string,
	text: string,
	idGenerator?: (kind: string) => string,
): Comment[] {
	const id = nextCommentId(idGenerator);
	return [...comments, { id, author, text, parentId }];
}

/** Comment ids anchored at the selection, outermost first (replies are not anchored). */
export function commentIdsAtSelection(view: EditorView): string[] {
	const { from, to, empty } = view.state.selection;
	const ids: string[] = [];
	view.state.doc.nodesBetween(
		empty ? Math.max(0, from - 1) : from,
		empty ? from + 1 : to,
		(node) => {
			const mark = node.marks.find((item) => item.type.name === 'comment');
			if (mark) for (const id of mark.attrs.ids as string[]) if (!ids.includes(id)) ids.push(id);
		},
	);
	return ids;
}

/** Where each comment's anchored text starts, in document order. */
export function commentAnchors(view: EditorView): Array<{ id: string; from: number }> {
	const seen = new Set<string>();
	const anchors: Array<{ id: string; from: number }> = [];
	view.state.doc.descendants((node, pos) => {
		const mark = node.marks.find((item) => item.type.name === 'comment');
		if (!mark) return;
		for (const id of mark.attrs.ids as string[])
			if (!seen.has(id)) {
				seen.add(id);
				anchors.push({ id, from: pos });
			}
	});
	return anchors;
}

/**
 * Review > Previous / Next Comment: moves the caret to the start of the neighbouring commented
 * text, wrapping around. Returns the comment id it landed on, or null when there are none.
 */
export function goToComment(view: EditorView, direction: 'previous' | 'next'): string | null {
	const anchors = commentAnchors(view);
	if (!anchors.length) return null;
	const here = view.state.selection.from;
	const target =
		direction === 'next'
			? (anchors.find((anchor) => anchor.from > here) ?? anchors[0]!)
			: ([...anchors].reverse().find((anchor) => anchor.from < here) ?? anchors.at(-1)!);
	view.dispatch(
		view.state.tr.setSelection(TextSelection.create(view.state.doc, target.from)).scrollIntoView(),
	);
	view.focus();
	return target.id;
}
