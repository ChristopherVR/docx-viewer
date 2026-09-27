// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Rebuilds a paragraph's inline content (runs, pictures, hyperlinks, tracked-change wrappers and
// comment anchors) from the model's flat run list. Combines the hyperlink grouping from
// write-hyperlink.ts with the revision/comment wrapping from write-revisions.ts.
import type { HyperlinkInfo, TextRun } from './model.js';
import type { RelationshipAllocator } from './relationship-allocator.js';
import { createRun } from './write-run.js';
import { hasSpecialBreak, isModeledBreak } from './breaks.js';
import { isCommentReferenceRun } from './parse-revisions.js';
import {
	convertToDeleteText,
	isCommentAnchorElement,
	isRevisionWrapperElement,
	revisionWrapper,
} from './write-revisions.js';
import { commentAnchorEdges, commentRangeEndNodes, commentRangeStartNode } from './write-comments.js';
import {
	getW,
	isElement,
	makeW,
	named,
	REL_NS,
	WORD_NS,
	type XmlDocument,
	type XmlElement,
} from './xml.js';

/** One modeled run's source element, with the hyperlink that wrapped it (if any). */
export interface InlineSlot {
	element: XmlElement;
	link?: XmlElement;
}

/** A run's own content is safe when it is plain text, a modeled break, or a lone picture. */
function runIsSafe(run: XmlElement): boolean {
	const content = Array.from(run.childNodes).filter(
		(child): child is XmlElement => isElement(child) && !named(child, 'rPr'),
	);
	if (content.length === 1 && named(content[0], 'br') && isModeledBreak(content[0])) return true;
	const kinds = new Set<string>();
	for (const element of content) {
		if (element.localName === 'br' && hasSpecialBreak(element)) return false;
		if (element.localName === 'cr' && element.attributes.length > 0) return false;
		const allowed = ['t', 'tab', 'br', 'cr', 'noBreakHyphen', 'delText', 'drawing', 'pict'];
		if (!allowed.includes(element.localName)) return false;
		kinds.add(element.localName);
	}
	// A run is modeled as text or as a picture, never both.
	return !((kinds.has('drawing') || kinds.has('pict')) && kinds.size > 1);
}

function runsOf(container: XmlElement, link: XmlElement | undefined, slots: InlineSlot[]): boolean {
	for (const child of Array.from(container.childNodes)) {
		if (!isElement(child)) {
			if (child.nodeType === 3 && child.textContent?.trim()) return false;
			continue;
		}
		if (named(child, 'r')) {
			if (isCommentReferenceRun(child)) continue;
			if (!runIsSafe(child)) return false;
			slots.push({ element: child, link });
		} else if (isRevisionWrapperElement(child) && !link) {
			if (!runsOf(child, undefined, slots)) return false;
		} else if (named(child, 'hyperlink') && !link) {
			if (!runsOf(child, child, slots)) return false;
		} else if (isCommentAnchorElement(child) && !link) continue;
		else return false;
	}
	return true;
}

/**
 * Collects the paragraph's modeled runs in the same order the parser produced them. Returns
 * `undefined` when the paragraph holds inline content this editor cannot safely regenerate.
 * `w:pPr` and bookmarks are handled by the caller.
 */
export function collectInlineSlots(paragraph: XmlElement): InlineSlot[] | undefined {
	const slots: InlineSlot[] = [];
	for (const child of Array.from(paragraph.childNodes)) {
		if (isElement(child) && (named(child, 'pPr') || named(child, 'bookmarkStart') || named(child, 'bookmarkEnd')))
			continue;
		if (!isElement(child)) {
			if (child.nodeType === 3 && child.textContent?.trim()) return undefined;
			continue;
		}
		if (named(child, 'r')) {
			if (isCommentReferenceRun(child)) continue;
			if (!runIsSafe(child)) return undefined;
			slots.push({ element: child });
		} else if (isRevisionWrapperElement(child)) {
			if (!runsOf(child, undefined, slots)) return undefined;
		} else if (named(child, 'hyperlink')) {
			if (!runsOf(child, child, slots)) return undefined;
		} else if (!isCommentAnchorElement(child)) return undefined;
	}
	return slots;
}

/** Top-level paragraph children that `buildInlineContent` output replaces. */
export function replaceableInlineChildren(paragraph: XmlElement): XmlElement[] {
	return Array.from(paragraph.childNodes).filter(
		(child): child is XmlElement =>
			isElement(child) &&
			(named(child, 'r') ||
				named(child, 'hyperlink') ||
				isRevisionWrapperElement(child) ||
				isCommentAnchorElement(child)),
	);
}

function sameLink(a: HyperlinkInfo | undefined, b: HyperlinkInfo | undefined): boolean {
	if (!a && !b) return true;
	if (!a || !b) return false;
	return a.href === b.href && a.anchor === b.anchor && a.tooltip === b.tooltip;
}

function setW(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}

function newHyperlink(doc: XmlDocument, link: HyperlinkInfo, allocator?: RelationshipAllocator): XmlElement {
	const container = makeW(doc, 'hyperlink');
	if (link.href) {
		if (!allocator) throw new Error('A relationship allocator is required to write new hyperlinks.');
		container.setAttributeNS(REL_NS, 'r:id', allocator.addExternalHyperlink(link.href));
	}
	if (link.anchor) setW(container, 'anchor', link.anchor);
	if (link.tooltip) setW(container, 'tooltip', link.tooltip);
	if (!getW(container, 'history')) setW(container, 'history', '1');
	return container;
}

/** Builds one run's nodes: its comment-range starts, the (revision-wrapped) run, and range ends. */
function runNodes(
	doc: XmlDocument,
	run: TextRun,
	base: TextRun | undefined,
	old: XmlElement | undefined,
	opens: string[],
	closes: string[],
	allocator?: RelationshipAllocator,
): XmlElement[] {
	const nodes = opens.map((id) => commentRangeStartNode(doc, id));
	const node = createRun(doc, run, base, old, allocator);
	const revision = run.revision;
	if (revision && ['insert', 'delete', 'moveFrom', 'moveTo'].includes(revision.kind)) {
		if (revision.kind === 'delete' || revision.kind === 'moveFrom') convertToDeleteText(doc, node);
		nodes.push(revisionWrapper(doc, revision, node));
	} else nodes.push(node);
	for (const id of closes) nodes.push(...commentRangeEndNodes(doc, id));
	return nodes;
}

/** Builds the paragraph's ordered top-level inline nodes from the model's flat run list. */
export function buildInlineContent(
	doc: XmlDocument,
	runs: TextRun[],
	base: TextRun[] | undefined,
	slots: InlineSlot[],
	allocator?: RelationshipAllocator,
): XmlElement[] {
	const { opens, closes } = commentAnchorEdges(runs);
	const nodesFor = (index: number) =>
		runNodes(
			doc,
			runs[index],
			base?.[index],
			slots[index]?.element,
			opens.get(index) ?? [],
			closes.get(index) ?? [],
			allocator,
		);
	const output: XmlElement[] = [];
	let index = 0;
	while (index < runs.length) {
		const link = runs[index].link;
		if (!link) {
			output.push(...nodesFor(index));
			index++;
			continue;
		}
		const start = index;
		while (index < runs.length && sameLink(runs[index].link, link)) index++;
		const oldContainers = new Set(slots.slice(start, index).map((slot) => slot?.link));
		const [onlyContainer] = oldContainers;
		const reusable =
			oldContainers.size === 1 &&
			onlyContainer !== undefined &&
			slots.length >= index &&
			base?.slice(start, index).every((run) => sameLink(run?.link, link)) === true;
		let container: XmlElement;
		if (reusable) {
			container = onlyContainer;
			for (const child of Array.from(container.childNodes)) container.removeChild(child);
		} else container = newHyperlink(doc, link, allocator);
		for (let item = start; item < index; item++)
			for (const node of nodesFor(item)) container.appendChild(node);
		output.push(container);
	}
	return output;
}
