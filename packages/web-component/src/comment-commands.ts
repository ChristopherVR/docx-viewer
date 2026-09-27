import type { Comment } from '@christophervr/docx-core';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { schema } from './schema';

let commentSerial = 0;
function nextCommentId(idGenerator?: (kind: string) => string): string {
	return idGenerator ? idGenerator('comment') : `dve-comment-${Date.now().toString(36)}-${++commentSerial}`;
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
		if (remaining.length) tr = tr.addMark(pos, pos + node.nodeSize, schema.marks.comment.create({ ids: remaining }));
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
