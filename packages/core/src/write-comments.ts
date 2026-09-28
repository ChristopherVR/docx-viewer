// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type JSZip from 'jszip';
import type { Comment } from './model.js';
import { buildXml, makeW, parseXml, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts.js';

const COMMENTS_PATH = 'word/comments.xml';
const COMMENTS_EXTENDED_PATH = 'word/commentsExtended.xml';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
/** Deterministic 8-hex-digit paragraph id, matching a comment to its own commentsExtended entry. */
function paraIdFor(id: string, index: number): string {
	let hash = (index + 1) >>> 0;
	for (let i = 0; i < id.length; i++) hash = (Math.imul(hash, 31) + id.charCodeAt(i)) >>> 0;
	return hash.toString(16).padStart(8, '0').toUpperCase();
}

function buildCommentsDocument(comments: Comment[]): XmlDocument {
	const doc = parseXml(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"></w:comments>`,
	);
	const root = doc.documentElement as XmlElement;
	comments.forEach((comment, index) => {
		const el = makeW(doc, 'comment');
		setAttribute(el, 'id', comment.id);
		setAttribute(el, 'author', comment.author);
		if (comment.date) setAttribute(el, 'date', comment.date);
		if (comment.initials) setAttribute(el, 'initials', comment.initials);
		const p = makeW(doc, 'p');
		p.setAttributeNS(W14_NS, 'w14:paraId', paraIdFor(comment.id, index));
		const r = makeW(doc, 'r');
		const t = makeW(doc, 't');
		t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
		t.appendChild(doc.createTextNode(comment.text));
		r.appendChild(t);
		p.appendChild(r);
		el.appendChild(p);
		root.appendChild(el);
	});
	return doc;
}

function buildCommentsExtendedDocument(comments: Comment[]): XmlDocument | null {
	if (!comments.some((comment) => comment.resolved || comment.parentId)) return null;
	const doc = parseXml(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="${W15_NS}"></w15:commentsEx>`,
	);
	const root = doc.documentElement as XmlElement;
	comments.forEach((comment, index) => {
		const entry = doc.createElementNS(W15_NS, 'w15:commentEx');
		entry.setAttributeNS(W15_NS, 'w15:paraId', paraIdFor(comment.id, index));
		if (comment.resolved) entry.setAttributeNS(W15_NS, 'w15:done', '1');
		const parentIndex = comments.findIndex((c) => c.id === comment.parentId);
		if (parentIndex >= 0)
			entry.setAttributeNS(W15_NS, 'w15:paraIdParent', paraIdFor(comment.parentId!, parentIndex));
		root.appendChild(entry);
	});
	return doc;
}

/**
 * Regenerates comments.xml/commentsExtended.xml from the model's comment list. Any comment text,
 * formatting or extended metadata not represented by `Comment` is not preserved once the list changes.
 */
export async function applyComments(zip: JSZip, comments: Comment[]): Promise<void> {
	const existed = Boolean(zip.file(COMMENTS_PATH));
	zip.file(COMMENTS_PATH, buildXml(buildCommentsDocument(comments)));
	const extended = buildCommentsExtendedDocument(comments);
	if (extended) zip.file(COMMENTS_EXTENDED_PATH, buildXml(extended));
	else zip.remove(COMMENTS_EXTENDED_PATH);
	if (!existed) {
		await ensureContentTypeOverride(
			zip,
			COMMENTS_PATH,
			'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
		);
		await ensureDocumentRelationship(
			zip,
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
			'comments.xml',
		);
	}
	if (extended) {
		await ensureContentTypeOverride(
			zip,
			COMMENTS_EXTENDED_PATH,
			'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
		);
		await ensureDocumentRelationship(
			zip,
			'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
			'commentsExtended.xml',
		);
	}
}

export function commentRangeStartNode(doc: XmlDocument, id: string): XmlElement {
	const element = makeW(doc, 'commentRangeStart');
	setAttribute(element, 'id', id);
	return element;
}
/** `commentRangeEnd` plus the anchoring `commentReference` run that follows it. */
export function commentRangeEndNodes(doc: XmlDocument, id: string): XmlElement[] {
	const end = makeW(doc, 'commentRangeEnd');
	setAttribute(end, 'id', id);
	const run = makeW(doc, 'r');
	const reference = makeW(doc, 'commentReference');
	setAttribute(reference, 'id', id);
	run.appendChild(reference);
	return [end, run];
}
