// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { expectDefined } from './expect-defined.js';
import type { Revision, TextRun } from './model.js';
import { children, first, getW, isElement, named, type XmlElement, WORD_NS } from './xml.js';

const REVISION_WRAPPERS: Record<string, Revision['kind']> = {
	ins: 'insert',
	del: 'delete',
	moveFrom: 'moveFrom',
	moveTo: 'moveTo',
};

/** Element local names this module owns inside a paragraph body; used by write.ts's safety checks. */
export const REVISION_WRAPPER_NAMES = Object.keys(REVISION_WRAPPERS);
export const COMMENT_ANCHOR_NAMES = ['commentRangeStart', 'commentRangeEnd', 'commentReference'];

/** Move ranges open at the current parse position (they may span paragraphs), innermost last. */
export interface OpenMoves {
	moveFrom: { name: string; id: string }[];
	moveTo: { name: string; id: string }[];
}
const MOVE_RANGES: Record<string, { kind: keyof OpenMoves; start: boolean }> = {
	moveFromRangeStart: { kind: 'moveFrom', start: true },
	moveFromRangeEnd: { kind: 'moveFrom', start: false },
	moveToRangeStart: { kind: 'moveTo', start: true },
	moveToRangeEnd: { kind: 'moveTo', start: false },
};

function trackMoveRange(item: XmlElement, moves: OpenMoves): boolean {
	if (item.namespaceURI && item.namespaceURI !== WORD_NS) return false;
	const range = MOVE_RANGES[item.localName ?? ''];
	if (!range) return false;
	const id = getW(item, 'id') ?? '';
	const open = moves[range.kind];
	if (range.start) open.push({ name: getW(item, 'name') ?? id, id });
	else {
		const index = open.findIndex((entry) => entry.id === id);
		if (index >= 0) open.splice(index, 1);
	}
	return true;
}

/** A `w:r` whose only content is a `w:commentReference`: a comment anchor, not document text. */
export function isCommentReferenceRun(run: XmlElement): boolean {
	const content = Array.from(run.childNodes).filter(
		(child): child is XmlElement => isElement(child) && !named(child, 'rPr'),
	);
	return content.length > 0 && content.every((child) => named(child, 'commentReference'));
}

function revisionFrom(node: XmlElement, kind: Revision['kind']): Revision {
	const revision: Revision = {
		kind,
		author: getW(node, 'author') || 'Unknown',
		id: getW(node, 'id') ?? '0',
	};
	const date = getW(node, 'date');
	if (date) revision.date = date;
	return revision;
}

/** Paragraph mark insertion/deletion, recorded as `w:pPr/w:rPr/w:ins|w:del`. */
export function paragraphMarkRevision(pPr: XmlElement | undefined): Revision | undefined {
	const rPr = first(pPr, 'rPr');
	const ins = first(rPr, 'ins');
	if (ins) return revisionFrom(ins, 'insert');
	const del = first(rPr, 'del');
	if (del) return revisionFrom(del, 'delete');
	return undefined;
}

/** Marks that `w:pPrChange` recorded a prior paragraph formatting snapshot (not itself modeled). */
export function paragraphFormatRevision(pPr: XmlElement | undefined): Revision | undefined {
	const change = first(pPr, 'pPrChange');
	return change ? revisionFrom(change, 'paragraphChange') : undefined;
}

/** Marks that `w:rPrChange` recorded a prior run formatting snapshot (not itself modeled). */
export function runFormatRevision(rPr: XmlElement | undefined): Revision | undefined {
	const change = first(rPr, 'rPrChange');
	return change ? revisionFrom(change, 'formatChange') : undefined;
}

/**
 * Walks a paragraph's direct children in document order, expanding `w:ins`/`w:del`/`w:moveFrom`/
 * `w:moveTo` wrappers into their contained runs (tagged with the wrapper's revision) and tracking
 * `w:commentRangeStart`/`w:commentRangeEnd` so runs inside an open range carry the comment's id.
 * Pass the same `active` array for consecutive paragraphs so a range spanning paragraphs tags
 * every run inside it.
 */
export function collectParagraphRuns(
	node: XmlElement,
	parseRun: (run: XmlElement, revision?: Revision) => TextRun,
	parseOther?: (item: XmlElement) => TextRun[] | undefined,
	resolveLink?: (hyperlink: XmlElement) => TextRun['link'],
	active: string[] = [],
	moves: OpenMoves = { moveFrom: [], moveTo: [] },
): { runs: TextRun[]; hasMove: boolean } {
	const runs: TextRun[] = [];
	let hasMove = false;
	const push = (run: TextRun) => {
		if (active.length) run.commentIds = [...active];
		runs.push(run);
	};
	for (const item of Array.from(node.childNodes).filter(isElement)) {
		if (named(item, 'r')) {
			if (!isCommentReferenceRun(item)) push(parseRun(item));
		} else if (named(item, 'hyperlink')) {
			const link = resolveLink?.(item);
			for (const run of children(item, 'r')) {
				const parsed = parseRun(run);
				if (link) parsed.link = link;
				push(parsed);
			}
		} else if (
			(!item.namespaceURI || item.namespaceURI === WORD_NS) &&
			Object.hasOwn(REVISION_WRAPPERS, item.localName)
		) {
			const kind = expectDefined(
				REVISION_WRAPPERS[item.localName],
				`revision wrapper ${item.localName}`,
			);
			if (kind === 'moveFrom' || kind === 'moveTo') hasMove = true;
			const revision = revisionFrom(item, kind);
			const range = kind === 'moveFrom' || kind === 'moveTo' ? moves[kind].at(-1) : undefined;
			if (range) revision.move = { name: range.name, rangeId: range.id };
			for (const run of children(item, 'r')) push(parseRun(run, revision));
		} else if (trackMoveRange(item, moves)) {
			continue;
		} else if (named(item, 'commentRangeStart')) {
			const id = getW(item, 'id');
			if (id !== undefined) active.push(id);
		} else if (named(item, 'commentRangeEnd')) {
			const id = getW(item, 'id');
			const index = id === undefined ? -1 : active.indexOf(id);
			if (index >= 0) active.splice(index, 1);
		} else {
			for (const run of parseOther?.(item) ?? []) push(run);
		}
	}
	return { runs, hasMove };
}
