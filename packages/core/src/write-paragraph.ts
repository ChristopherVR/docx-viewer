// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Paragraph } from './model.js';
import { children, first, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';
import { writeParagraphProperties } from './write-paragraph-properties.js';
import { orderParagraphProperties, writeTabStops } from './tab-stops.js';
import { continuationKey } from './comment-spans.js';
import type { CommentSpans } from './write-ranges.js';
import { reconcileBookmarks } from './bookmarks.js';
import { writeNumberingProperties } from './numbering-write.js';
import { writeParagraphMarkRevision } from './write-revisions.js';
import { paragraphJustification } from './paragraph-alignment.js';
import { runHasUnknownProperties } from './write-run-validation.js';
import {
	buildInlineContent,
	collectInlineSlots,
	replaceableInlineChildren,
} from './write-inline.js';
import type { RelationshipAllocator } from './relationship-allocator.js';

export function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
}

function rejectUnsafeRunSegmentation(
	paragraph: Paragraph,
	base: Paragraph | undefined,
	oldRuns: XmlElement[],
): void {
	if (!base || !oldRuns.some(runHasUnknownProperties)) return;
	const sameText =
		paragraph.runs.map((run) => run.text).join('') === base.runs.map((run) => run.text).join('');
	const sameBoundaries =
		paragraph.runs.length === base.runs.length &&
		paragraph.runs.every((run, index) => run.text === base.runs[index]?.text);
	if (sameBoundaries || (!sameText && paragraph.runs.length === base.runs.length)) return;
	throw new Error(
		`Cannot edit paragraph ${paragraph.id}: changing run boundaries could drop unsupported run properties. The original DOCX package remains unchanged.`,
	);
}

export function writeParagraphImpl(
	doc: XmlDocument,
	paragraph: Paragraph,
	node: XmlElement,
	base: Paragraph | undefined,
	allocator: RelationshipAllocator | undefined,
	spans?: CommentSpans,
): XmlElement {
	const continuation = spans?.next.get(paragraph.id);
	const ranges = spans && {
		...(continuation ?? { before: new Set<string>(), after: new Set<string>() }),
		rangeIds: spans.rangeIds,
	};
	const sameRanges =
		continuationKey(continuation) === continuationKey(spans?.original.get(paragraph.id));
	if (base && sameRanges && JSON.stringify(paragraph) === JSON.stringify(base)) return node;
	const slots = collectInlineSlots(node);
	if (!slots)
		throw new Error(
			`Cannot edit paragraph ${paragraph.id}: it contains inline OOXML that this editor cannot safely relocate. The original DOCX package remains unchanged.`,
		);
	let pPr = first(node, 'pPr');
	if (!pPr) {
		pPr = makeW(doc, 'pPr');
		node.insertBefore(pPr, node.firstChild);
	}
	if (
		!base ||
		paragraph.align !== base.align ||
		paragraph.justification !== base.justification ||
		paragraph.direction !== base.direction
	) {
		removeChildren(pPr, 'jc');
		const jc = paragraphJustification(paragraph);
		if (jc) {
			const align = makeW(doc, 'jc');
			setAttribute(align, 'val', jc);
			pPr.appendChild(align);
		}
	}
	if (!base || paragraph.style !== base.style) {
		removeChildren(pPr, 'pStyle');
		if (paragraph.style) {
			const style = makeW(doc, 'pStyle');
			setAttribute(style, 'val', paragraph.style);
			pPr.appendChild(style);
		}
	}
	writeParagraphProperties(doc, pPr, paragraph, base);
	writeNumberingProperties(doc, pPr, paragraph, base);
	writeParagraphMarkRevision(doc, pPr, paragraph, base);
	removeChildren(pPr, 'pPrChange');
	writeTabStops(doc, pPr, paragraph, base);
	orderParagraphProperties(pPr);
	rejectUnsafeRunSegmentation(
		paragraph,
		base,
		slots.flatMap((slot) => (slot.element ? [slot.element] : [])),
	);
	// Bookmarks are preserved but not repositioned precisely: an edited paragraph's bookmarks move
	// to its boundaries (starts right after pPr, ends at the close) instead of their exact original
	// run offsets, which this run-level model does not track.
	const { starts: bookmarkStarts, ends: bookmarkEnds } = reconcileBookmarks(
		doc,
		children(node, 'bookmarkStart'),
		children(node, 'bookmarkEnd'),
		paragraph.bookmarks ?? [],
		base?.bookmarks ?? [],
		() => spans?.nextId() ?? '0',
	);
	for (const bookmark of [...children(node, 'bookmarkStart'), ...children(node, 'bookmarkEnd')])
		node.removeChild(bookmark);
	for (const oldNode of replaceableInlineChildren(node)) node.removeChild(oldNode);
	const newNodes = buildInlineContent(doc, paragraph.runs, base?.runs, slots, allocator, ranges);
	let anchor: XmlElement = pPr;
	for (const bookmark of bookmarkStarts) {
		node.insertBefore(bookmark, anchor.nextSibling);
		anchor = bookmark;
	}
	for (const run of newNodes) {
		node.insertBefore(run, anchor.nextSibling);
		anchor = run;
	}
	for (const bookmark of bookmarkEnds) {
		node.insertBefore(bookmark, anchor.nextSibling);
		anchor = bookmark;
	}
	if (!pPr.childNodes.length) node.removeChild(pPr);
	return node;
}

export function createParagraph(
	doc: XmlDocument,
	paragraph: Paragraph,
	allocator: RelationshipAllocator | undefined,
	spans?: CommentSpans,
): XmlElement {
	const node = makeW(doc, 'p');
	return writeParagraphImpl(doc, paragraph, node, undefined, allocator, spans);
}
