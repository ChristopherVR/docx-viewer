// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Comment } from './model.js';
import { getW, parseXml, textContent, type XmlElement, WORD_NS } from './xml.js';

const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

function paragraphText(paragraph: XmlElement): string {
	return Array.from(paragraph.getElementsByTagNameNS(WORD_NS, 't'))
		.map((node) => textContent(node as XmlElement))
		.join('');
}

/** Parses comments.xml and, when present, commentsExtended.xml's done/reply metadata. */
export function parseComments(commentsXml: string, extendedXml?: string): Comment[] {
	const doc = parseXml(commentsXml);
	const commentElements = Array.from(
		doc.getElementsByTagNameNS(WORD_NS, 'comment'),
	) as XmlElement[];
	const paraIdToCommentId = new Map<string, string>();
	const comments: Comment[] = commentElements.map((element) => {
		const id = getW(element, 'id') ?? '0';
		const text = Array.from(element.getElementsByTagNameNS(WORD_NS, 'p'))
			.map((p) => paragraphText(p as XmlElement))
			.join('\n')
			.trim();
		const comment: Comment = { id, author: getW(element, 'author') || 'Unknown', text };
		const initials = getW(element, 'initials');
		if (initials) comment.initials = initials;
		const date = getW(element, 'date');
		if (date) comment.date = date;
		const firstParagraph = element.getElementsByTagNameNS(WORD_NS, 'p')[0] as XmlElement | undefined;
		const paraId = firstParagraph?.getAttributeNS(W14_NS, 'paraId') || undefined;
		if (paraId) paraIdToCommentId.set(paraId, id);
		return comment;
	});
	if (extendedXml) {
		const extended = parseXml(extendedXml);
		for (const entry of Array.from(
			extended.getElementsByTagNameNS(W15_NS, 'commentEx'),
		) as XmlElement[]) {
			const paraId = entry.getAttributeNS(W15_NS, 'paraId');
			const comment = paraId && comments.find((c) => c.id === paraIdToCommentId.get(paraId));
			if (!comment) continue;
			const done = entry.getAttributeNS(W15_NS, 'done');
			comment.resolved = done === '1' || done === 'true';
			const parentParaId = entry.getAttributeNS(W15_NS, 'paraIdParent');
			const parentId = parentParaId && paraIdToCommentId.get(parentParaId);
			if (parentId) comment.parentId = parentId;
		}
	}
	return comments;
}
