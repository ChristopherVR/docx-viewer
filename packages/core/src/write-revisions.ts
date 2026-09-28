// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Paragraph, Revision } from './model.js';
import { children, first, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';

const REVISION_WRAPPER_NAMES = ['ins', 'del', 'moveFrom', 'moveTo'];
/** Range markers the writer regenerates from the model: comment anchors and move ranges. */
const COMMENT_ANCHOR_NAMES = [
	'commentRangeStart',
	'commentRangeEnd',
	'commentReference',
	'moveFromRangeStart',
	'moveFromRangeEnd',
	'moveToRangeStart',
	'moveToRangeEnd',
];

function isWordElement(element: XmlElement, names: string[]): boolean {
	return (
		(!element.namespaceURI || element.namespaceURI === WORD_NS) && names.includes(element.localName)
	);
}
export const isRevisionWrapperElement = (element: XmlElement): boolean =>
	isWordElement(element, REVISION_WRAPPER_NAMES);
export const isCommentAnchorElement = (element: XmlElement): boolean =>
	isWordElement(element, COMMENT_ANCHOR_NAMES);

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
