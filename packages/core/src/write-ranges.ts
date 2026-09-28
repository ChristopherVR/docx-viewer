// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Comment and move range markers around rewritten runs. A range may span paragraphs, so each
// paragraph only opens the ranges that start in it and closes the ones that end in it.
import type { Block, Paragraph, TextRun } from './model.js';
import { rangeKeys, type CommentContinuation } from './comment-spans.js';
import { commentRangeEndNodes, commentRangeStartNode } from './write-comments.js';
import { getW, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';

function setW(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}

/** How this paragraph's comment and move ranges continue across paragraphs, plus new range ids. */
export interface ParagraphRanges extends CommentContinuation {
	/** `w:id`s for move ranges that have none yet (new moves), by range key. */
	rangeIds: ReadonlyMap<string, string>;
}

/** Comment range continuations in the blocks being written and in the part as loaded. */
export interface CommentSpans {
	next: Map<string, CommentContinuation>;
	original: Map<string, CommentContinuation>;
	/** Fresh `w:id`s for move ranges created in the editor (their runs carry no range id yet). */
	rangeIds: Map<string, string>;
	/** Allocates `w:id`s for new annotations such as bookmarks. */
	nextId: () => string;
}

/**
 * Hands out annotation `w:id`s (bookmarks, move ranges) after the largest id already in the part
 * or among the model's revisions, so new markers never collide with existing ones.
 */
export function annotationIdAllocator(doc: XmlDocument, blocks: Block[]): () => string {
	let next: number | undefined;
	return () => {
		if (next === undefined) {
			let max = -1;
			const reserve = (id: string | undefined) => {
				const value = Number(id);
				if (id && Number.isSafeInteger(value) && value > max) max = value;
			};
			forEachParagraph(blocks, (paragraph) => {
				reserve(paragraph.markRevision?.id);
				for (const run of paragraph.runs) reserve(run.revision?.id);
			});
			for (const element of Array.from(doc.getElementsByTagNameNS(WORD_NS, '*')))
				reserve(getW(element, 'id'));
			next = max + 1;
		}
		return String(next++);
	};
}

function forEachParagraph(blocks: Block[], visit: (paragraph: Paragraph) => void): void {
	for (const block of blocks)
		if (block.type === 'paragraph') visit(block);
		else for (const row of block.rows) for (const cell of row) cell.paragraphs.forEach(visit);
}

/** `w:id`s for move ranges created in the editor, whose runs carry no range id yet. */
export function newMoveRangeIds(blocks: Block[], nextId: () => string): Map<string, string> {
	const ids = new Map<string, string>();
	forEachParagraph(blocks, (paragraph) => {
		for (const run of paragraph.runs) {
			const move = run.revision?.move;
			const key = move && `${run.revision!.kind}:${move.name}`;
			if (key && !move.rangeId && !ids.has(key)) ids.set(key, nextId());
		}
	});
	return ids;
}

/** A range marker to emit around a run: a comment anchor or a move range. */
export type RangeEdge =
	| { kind: 'comment'; id: string }
	| { kind: 'moveFrom' | 'moveTo'; name: string; rangeId: string; revision: TextRun['revision'] };

export function rangeStartNodes(doc: XmlDocument, edge: RangeEdge): XmlElement[] {
	if (edge.kind === 'comment') return [commentRangeStartNode(doc, edge.id)];
	const start = makeW(doc, `${edge.kind}RangeStart`);
	setW(start, 'id', edge.rangeId);
	if (edge.revision?.author) setW(start, 'author', edge.revision.author);
	if (edge.revision?.date) setW(start, 'date', edge.revision.date);
	setW(start, 'name', edge.name);
	return [start];
}
export function rangeEndNodes(doc: XmlDocument, edge: RangeEdge): XmlElement[] {
	if (edge.kind === 'comment') return commentRangeEndNodes(doc, edge.id);
	const end = makeW(doc, `${edge.kind}RangeEnd`);
	setW(end, 'id', edge.rangeId);
	return [end];
}

/** Range edges per run index: where each comment or move range opens and closes in this paragraph. */
export function rangeEdges(
	runs: TextRun[],
	ranges: ParagraphRanges | undefined,
): { opens: Map<number, RangeEdge[]>; closes: Map<number, RangeEdge[]> } {
	const first = new Map<string, number>();
	const last = new Map<string, number>();
	runs.forEach((run, index) => {
		for (const key of rangeKeys(run)) {
			if (!first.has(key)) first.set(key, index);
			last.set(key, index);
		}
	});
	const edgeFor = (key: string, index: number): RangeEdge => {
		const separator = key.indexOf(':');
		const kind = key.slice(0, separator);
		const name = key.slice(separator + 1);
		if (kind === 'comment') return { kind: 'comment', id: name };
		const revision = runs[index].revision;
		return {
			kind: kind as 'moveFrom' | 'moveTo',
			name,
			rangeId: revision?.move?.rangeId ?? ranges?.rangeIds.get(key) ?? '0',
			revision,
		};
	};
	const opens = new Map<number, RangeEdge[]>();
	const closes = new Map<number, RangeEdge[]>();
	// Comments sort before moves when opening, so a comment encloses a move range it overlaps.
	const order = (key: string) => (key.startsWith('comment:') ? 0 : 1);
	for (const key of [...first.keys()].sort((a, b) => order(a) - order(b))) {
		if (!ranges?.before.has(key)) {
			const index = first.get(key)!;
			(opens.get(index) ?? opens.set(index, []).get(index)!).push(edgeFor(key, index));
		}
		if (!ranges?.after.has(key)) {
			const index = last.get(key)!;
			(closes.get(index) ?? closes.set(index, []).get(index)!).push(edgeFor(key, index));
		}
	}
	return { opens, closes };
}
