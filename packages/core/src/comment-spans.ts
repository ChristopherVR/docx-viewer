// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Ranges that span paragraphs: each comment gets one w:commentRangeStart in the first paragraph it
// touches and one w:commentRangeEnd (plus its reference) in the last; move ranges
// (w:moveFromRangeStart/w:moveToRangeStart) work the same way.
import type { Block, DocumentModel, Paragraph, TextRun } from './model.js';

/** Range keys a run belongs to: `comment:<id>` and, for moved text, `moveFrom:<name>`/`moveTo:<name>`. */
export function rangeKeys(run: TextRun): string[] {
	const keys = (run.commentIds ?? []).map((id) => `comment:${id}`);
	const revision = run.revision;
	if (revision?.move && (revision.kind === 'moveFrom' || revision.kind === 'moveTo'))
		keys.push(`${revision.kind}:${revision.move.name}`);
	return keys;
}

/** Range keys on a paragraph that started in an earlier paragraph or continue into a later one. */
export interface CommentContinuation {
	before: Set<string>;
	after: Set<string>;
}

function paragraphsOf(blocks: Block[]): Paragraph[] {
	return blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
	);
}

/** Per paragraph id, the comment and move ranges crossing its start or end, in document order. */
export function commentContinuations(blocks: Block[]): Map<string, CommentContinuation> {
	const paragraphs = paragraphsOf(blocks);
	const first = new Map<string, number>();
	const last = new Map<string, number>();
	const idsByParagraph = paragraphs.map((paragraph, index) => {
		const ids = new Set(paragraph.runs.flatMap(rangeKeys));
		for (const id of ids) {
			if (!first.has(id)) first.set(id, index);
			last.set(id, index);
		}
		return ids;
	});
	const result = new Map<string, CommentContinuation>();
	paragraphs.forEach((paragraph, index) => {
		const ids = [...(idsByParagraph[index] ?? [])];
		const before = new Set(ids.filter((id) => (first.get(id) ?? index) < index));
		const after = new Set(ids.filter((id) => (last.get(id) ?? index) > index));
		if (before.size || after.size) result.set(paragraph.id, { before, after });
	});
	return result;
}

/** A comparable form, to tell when a paragraph's range boundaries changed without its content. */
export function continuationKey(continuation: CommentContinuation | undefined): string {
	if (!continuation) return '';
	return JSON.stringify([[...continuation.before].sort(), [...continuation.after].sort()]);
}

const DECIMAL_ID = /^\d{1,9}$/;

/**
 * Word requires decimal comment ids (`ST_DecimalNumber`). Editors mint readable ids such as
 * `dve-comment-…`; this renumbers those (and duplicates) to unused decimals everywhere they appear:
 * the comment list, reply parents and run anchors. Returns the model unchanged when all are valid.
 */
export function numberCommentIds(model: DocumentModel): DocumentModel {
	const comments = model.comments ?? [];
	const valid = new Set(comments.map((comment) => comment.id).filter((id) => DECIMAL_ID.test(id)));
	if (valid.size === comments.length) return model;
	let next = Math.max(-1, ...[...valid].map(Number)) + 1;
	const renamed = new Map<string, string>();
	for (const comment of comments)
		if (!DECIMAL_ID.test(comment.id) && !renamed.has(comment.id))
			renamed.set(comment.id, String(next++));
	const rename = (id: string) => renamed.get(id) ?? id;
	const paragraph = (block: Paragraph): Paragraph => ({
		...block,
		runs: block.runs.map((run) =>
			run.commentIds ? { ...run, commentIds: run.commentIds.map(rename) } : run,
		),
	});
	const blocks = model.blocks.map((block) =>
		block.type === 'paragraph'
			? paragraph(block)
			: {
					...block,
					rows: block.rows.map((row) =>
						row.map((cell) => ({ ...cell, paragraphs: cell.paragraphs.map(paragraph) })),
					),
				},
	);
	return {
		...model,
		blocks,
		comments: comments.map((comment) => ({
			...comment,
			id: rename(comment.id),
			...(comment.parentId ? { parentId: rename(comment.parentId) } : {}),
		})),
	};
}
