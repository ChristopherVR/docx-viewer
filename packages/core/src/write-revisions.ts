// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Paragraph, Revision, TextRun } from './model.js';
import {
	children,
	first,
	isElement,
	makeW,
	named,
	type XmlDocument,
	type XmlElement,
	WORD_NS,
} from './xml.js';
import { createRun } from './write-run.js';
import {
	commentAnchorEdges,
	commentRangeEndNodes,
	commentRangeStartNode,
} from './write-comments.js';

const REVISION_WRAPPER_NAMES = ['ins', 'del', 'moveFrom', 'moveTo'];
const COMMENT_ANCHOR_NAMES = ['commentRangeStart', 'commentRangeEnd', 'commentReference'];

function isWordElement(element: XmlElement, names: string[]): boolean {
	return (
		(!element.namespaceURI || element.namespaceURI === WORD_NS) && names.includes(element.localName)
	);
}
export const isRevisionWrapperElement = (element: XmlElement): boolean =>
	isWordElement(element, REVISION_WRAPPER_NAMES);
export const isCommentAnchorElement = (element: XmlElement): boolean =>
	isWordElement(element, COMMENT_ANCHOR_NAMES);
export const isManagedParagraphChild = (element: XmlElement): boolean =>
	isRevisionWrapperElement(element) || isCommentAnchorElement(element);

/** Collects `<w:r>` elements in document order, expanding ins/del/moveFrom/moveTo wrappers. */
export function gatherOldRuns(node: XmlElement): XmlElement[] {
	const runs: XmlElement[] = [];
	for (const child of Array.from(node.childNodes).filter(isElement)) {
		if (named(child, 'r')) runs.push(child);
		else if (isRevisionWrapperElement(child)) runs.push(...children(child, 'r'));
	}
	return runs;
}

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
}
function revisionTag(kind: Revision['kind']): string {
	return kind === 'delete'
		? 'del'
		: kind === 'moveFrom'
			? 'moveFrom'
			: kind === 'moveTo'
				? 'moveTo'
				: 'ins';
}
export function revisionWrapper(
	doc: XmlDocument,
	revision: Revision,
	content: XmlElement,
): XmlElement {
	const wrapper = makeW(doc, revisionTag(revision.kind));
	setAttribute(wrapper, 'id', revision.id);
	setAttribute(wrapper, 'author', revision.author);
	if (revision.date) setAttribute(wrapper, 'date', revision.date);
	wrapper.appendChild(content);
	return wrapper;
}
/** Deleted/moved-from text is stored as `w:delText` rather than `w:t`. */
export function convertToDeleteText(doc: XmlDocument, run: XmlElement): void {
	for (const t of children(run, 't')) {
		const delText = doc.createElementNS(WORD_NS, 'w:delText');
		const space = t.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'space');
		if (space) delText.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', space);
		while (t.firstChild) delText.appendChild(t.firstChild);
		run.replaceChild(delText, t);
	}
}

/**
 * Builds the flat list of paragraph-body nodes — comment range anchors, ins/del/moveFrom/moveTo
 * wrappers, and plain runs — for the paragraph's current run list.
 */
export function buildInlineNodes(
	doc: XmlDocument,
	runs: TextRun[],
	baseRuns: TextRun[] | undefined,
	oldRuns: XmlElement[],
): XmlElement[] {
	const { opens, closes } = commentAnchorEdges(runs);
	const output: XmlElement[] = [];
	runs.forEach((run, index) => {
		for (const id of opens.get(index) ?? []) output.push(commentRangeStartNode(doc, id));
		const runNode = createRun(doc, run, baseRuns?.[index], oldRuns[index]);
		const revision = run.revision;
		if (
			revision &&
			(revision.kind === 'insert' ||
				revision.kind === 'delete' ||
				revision.kind === 'moveFrom' ||
				revision.kind === 'moveTo')
		) {
			if (revision.kind === 'delete' || revision.kind === 'moveFrom')
				convertToDeleteText(doc, runNode);
			output.push(revisionWrapper(doc, revision, runNode));
		} else output.push(runNode);
		for (const id of closes.get(index) ?? []) output.push(...commentRangeEndNodes(doc, id));
	});
	return output;
}

/** Writes/clears the paragraph mark's own tracked insertion/deletion (`w:pPr/w:rPr/w:ins|w:del`). */
export function writeParagraphMarkRevision(
	doc: XmlDocument,
	pPr: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	if (base && JSON.stringify(paragraph.markRevision) === JSON.stringify(base.markRevision)) return;
	let rPr = first(pPr, 'rPr');
	if (rPr) {
		removeChildren(rPr, 'ins');
		removeChildren(rPr, 'del');
	}
	if (!paragraph.markRevision) {
		if (rPr && !rPr.attributes.length && !rPr.childNodes.length) pPr.removeChild(rPr);
		return;
	}
	if (!rPr) {
		rPr = makeW(doc, 'rPr');
		pPr.appendChild(rPr);
	}
	const revision = paragraph.markRevision;
	const element = makeW(doc, revision.kind === 'delete' ? 'del' : 'ins');
	setAttribute(element, 'id', revision.id);
	setAttribute(element, 'author', revision.author);
	if (revision.date) setAttribute(element, 'date', revision.date);
	rPr.insertBefore(element, rPr.firstChild);
}
