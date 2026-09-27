// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
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
 * Comment ranges are assumed to stay within one paragraph; multi-paragraph ranges are not modeled.
 */
export function collectParagraphRuns(
	node: XmlElement,
	parseRun: (run: XmlElement, revision?: Revision) => TextRun,
	parseOther?: (item: XmlElement) => TextRun | undefined,
): { runs: TextRun[]; hasMove: boolean } {
	const runs: TextRun[] = [];
	const active: string[] = [];
	let hasMove = false;
	const push = (run: TextRun) => {
		if (active.length) run.commentIds = [...active];
		runs.push(run);
	};
	for (const item of Array.from(node.childNodes).filter(isElement)) {
		if (named(item, 'r')) push(parseRun(item));
		else if (named(item, 'hyperlink'))
			for (const run of children(item, 'r')) push(parseRun(run));
		else if (
			(!item.namespaceURI || item.namespaceURI === WORD_NS) &&
			Object.hasOwn(REVISION_WRAPPERS, item.localName)
		) {
			const kind = REVISION_WRAPPERS[item.localName];
			if (kind === 'moveFrom' || kind === 'moveTo') hasMove = true;
			const revision = revisionFrom(item, kind);
			for (const run of children(item, 'r')) push(parseRun(run, revision));
		} else if (named(item, 'commentRangeStart')) {
			const id = getW(item, 'id');
			if (id !== undefined) active.push(id);
		} else if (named(item, 'commentRangeEnd')) {
			const id = getW(item, 'id');
			const index = id === undefined ? -1 : active.indexOf(id);
			if (index >= 0) active.splice(index, 1);
		} else {
			const run = parseOther?.(item);
			if (run) push(run);
		}
	}
	return { runs, hasMove };
}
