// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Rebuilds a paragraph's inline content (runs, pictures, hyperlinks, tracked-change wrappers and
// comment/move range markers) from the model's flat run list: hyperlink and field grouping plus
// the revision wrappers from write-revisions.ts and range markers from write-ranges.ts.
import type { HyperlinkInfo, TextRun } from './model.js';
import type { RelationshipAllocator } from './relationship-allocator.js';
import {
	rangeEdges,
	rangeEndNodes,
	rangeStartNodes,
	type ParagraphRanges,
	type RangeEdge,
} from './write-ranges.js';

import { createRun } from './write-run.js';
import { hasSpecialBreak, isModeledBreak } from './breaks.js';
import { isCommentReferenceRun } from './parse-revisions.js';
import {
	convertToDeleteText,
	isCommentAnchorElement,
	isRevisionWrapperElement,
	revisionWrapper,
} from './write-revisions.js';
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

/**
 * One modeled run's source element, with the hyperlink or simple field that wrapped it (if any).
 * A simple field without a cached result has a slot with no element.
 */
export interface InlineSlot {
	element?: XmlElement;
	container?: XmlElement;
}

/** A run's own content is safe when it is plain text, a modeled break, or a lone picture. */
function runIsSafe(run: XmlElement): boolean {
	const content = Array.from(run.childNodes).filter(
		(child): child is XmlElement => isElement(child) && !named(child, 'rPr'),
	);
	if (content.length === 1 && named(content[0], 'br') && isModeledBreak(content[0])) return true;
	// Note references and a note's own number mark are modeled runs (see block-parser.ts).
	if (
		content.length === 1 &&
		['footnoteReference', 'endnoteReference', 'footnoteRef', 'endnoteRef'].some((name) =>
			named(content[0], name),
		)
	)
		return true;
	// A complex field's marker (without form-field data) or its instruction text are modeled runs.
	if (content.length === 1 && named(content[0], 'fldChar') && !content[0].childNodes.length)
		return ['begin', 'separate', 'end'].includes(getW(content[0], 'fldCharType') ?? '');
	if (content.length && content.every((element) => named(element, 'instrText'))) return true;
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

function runsOf(
	parent: XmlElement,
	container: XmlElement | undefined,
	slots: InlineSlot[],
): boolean {
	for (const child of Array.from(parent.childNodes)) {
		if (!isElement(child)) {
			if (child.nodeType === 3 && child.textContent?.trim()) return false;
			continue;
		}
		if (named(child, 'r')) {
			if (isCommentReferenceRun(child)) continue;
			if (!runIsSafe(child)) return false;
			slots.push({ element: child, ...(container && { container }) });
		} else if (isRevisionWrapperElement(child) && !container) {
			if (!runsOf(child, undefined, slots)) return false;
		} else if ((named(child, 'hyperlink') || named(child, 'fldSimple')) && !container) {
			const before = slots.length;
			if (!runsOf(child, child, slots)) return false;
			// The parser shows a placeholder for a simple field without a cached result.
			if (named(child, 'fldSimple') && slots.length === before) slots.push({ container: child });
		} else if (isCommentAnchorElement(child) && !container) continue;
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
		if (
			isElement(child) &&
			(named(child, 'pPr') || named(child, 'bookmarkStart') || named(child, 'bookmarkEnd'))
		)
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
		} else if (named(child, 'hyperlink') || named(child, 'fldSimple')) {
			const before = slots.length;
			if (!runsOf(child, child, slots)) return undefined;
			if (named(child, 'fldSimple') && slots.length === before) slots.push({ container: child });
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
				named(child, 'fldSimple') ||
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

function newHyperlink(
	doc: XmlDocument,
	link: HyperlinkInfo,
	allocator?: RelationshipAllocator,
): XmlElement {
	const container = makeW(doc, 'hyperlink');
	if (link.href) {
		if (!allocator)
			throw new Error('A relationship allocator is required to write new hyperlinks.');
		container.setAttributeNS(REL_NS, 'r:id', allocator.addExternalHyperlink(link.href));
	}
	if (link.anchor) setW(container, 'anchor', link.anchor);
	if (link.tooltip) setW(container, 'tooltip', link.tooltip);
	if (!getW(container, 'history')) setW(container, 'history', '1');
	return container;
}

/** Builds one run's nodes: its range starts, the (revision-wrapped) run, and range ends. */
function runNodes(
	doc: XmlDocument,
	run: TextRun,
	base: TextRun | undefined,
	old: XmlElement | undefined,
	opens: RangeEdge[],
	closes: RangeEdge[],
	allocator?: RelationshipAllocator,
): XmlElement[] {
	// Comments open outside moves and close outside them, so ranges nest.
	const nodes = opens.flatMap((edge) => rangeStartNodes(doc, edge));
	const node = createRun(doc, run, base, old, allocator);
	const revision = run.revision;
	if (revision && ['insert', 'delete', 'moveFrom', 'moveTo'].includes(revision.kind)) {
		if (revision.kind === 'delete' || revision.kind === 'moveFrom') convertToDeleteText(doc, node);
		nodes.push(revisionWrapper(doc, revision, node));
	} else nodes.push(node);
	for (const edge of [...closes].reverse()) nodes.push(...rangeEndNodes(doc, edge));
	return nodes;
}

/** Builds the paragraph's ordered top-level inline nodes from the model's flat run list. */
export function buildInlineContent(
	doc: XmlDocument,
	runs: TextRun[],
	base: TextRun[] | undefined,
	slots: InlineSlot[],
	allocator?: RelationshipAllocator,
	ranges?: ParagraphRanges,
): XmlElement[] {
	// Ranges continuing from an earlier paragraph or into a later one open or close there instead.
	const { opens, closes } = rangeEdges(runs, ranges);
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
		const group = containerGroup(runs[index]);
		if (!group) {
			output.push(...nodesFor(index));
			index++;
			continue;
		}
		const start = index;
		while (index < runs.length && sameGroup(containerGroup(runs[index]), group)) index++;
		const oldContainers = new Set(slots.slice(start, index).map((slot) => slot?.container));
		const [onlyContainer] = oldContainers;
		const reusable =
			oldContainers.size === 1 &&
			onlyContainer !== undefined &&
			onlyContainer.localName === (group.kind === 'link' ? 'hyperlink' : 'fldSimple') &&
			base?.slice(start, index).every((run) => sameGroup(run && containerGroup(run), group)) ===
				true;
		let container: XmlElement;
		if (reusable) {
			container = onlyContainer;
			for (const child of Array.from(container.childNodes)) container.removeChild(child);
		} else if (group.kind === 'link') container = newHyperlink(doc, group.link, allocator);
		else {
			container = makeW(doc, 'fldSimple');
			setW(container, 'instr', ` ${group.instr} `);
		}
		for (let item = start; item < index; item++)
			for (const node of nodesFor(item)) container.appendChild(node);
		output.push(container);
	}
	return output;
}

type ContainerGroup =
	| { kind: 'link'; link: HyperlinkInfo }
	| { kind: 'simpleField'; instr: string };

/** The wrapper a run is written inside: a hyperlink, or a simple field (`w:fldSimple`). */
function containerGroup(run: TextRun): ContainerGroup | undefined {
	if (run.link) return { kind: 'link', link: run.link };
	if (run.field?.simple) return { kind: 'simpleField', instr: run.field.instr };
	return undefined;
}

function sameGroup(a: ContainerGroup | undefined, b: ContainerGroup | undefined): boolean {
	if (!a || !b) return a === b;
	if (a.kind === 'link' && b.kind === 'link') return sameLink(a.link, b.link);
	return a.kind === 'simpleField' && b.kind === 'simpleField' && a.instr === b.instr;
}
