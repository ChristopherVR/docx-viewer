// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Groups consecutive same-target runs into w:hyperlink wrappers, reusing the original element
// (and its relationship) verbatim when the link is unchanged, and allocating a new relationship
// only when a link is newly created or its target changes.
import type { HyperlinkInfo, TextRun } from './model.js';
import type { RelationshipAllocator } from './relationship-allocator.js';
import { createRun } from './write-run.js';
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

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}

export type RunSlot =
	| { kind: 'plain'; element: XmlElement }
	| { kind: 'link'; element: XmlElement; container: XmlElement };

/**
 * Collects the paragraph's run-bearing children in document order (plain runs, and runs nested
 * in a `w:hyperlink`). Returns `undefined` when the paragraph has other top-level content this
 * editor cannot safely rewrite (bookmarks are handled separately by the caller).
 */
export function collectRunSlots(node: XmlElement): RunSlot[] | undefined {
	const slots: RunSlot[] = [];
	for (const child of Array.from(node.childNodes)) {
		if (!isElement(child)) {
			if (child.nodeType === 3 && child.textContent?.trim()) return undefined;
			continue;
		}
		if (named(child, 'pPr') || named(child, 'bookmarkStart') || named(child, 'bookmarkEnd'))
			continue;
		if (named(child, 'r')) {
			slots.push({ kind: 'plain', element: child });
			continue;
		}
		if (named(child, 'hyperlink')) {
			for (const run of Array.from(child.childNodes)) {
				if (!isElement(run)) {
					if (run.nodeType === 3 && run.textContent?.trim()) return undefined;
					continue;
				}
				if (!named(run, 'r')) return undefined;
				slots.push({ kind: 'link', element: run, container: child });
			}
			continue;
		}
		return undefined;
	}
	return slots;
}

function sameLink(a: HyperlinkInfo | undefined, b: HyperlinkInfo | undefined): boolean {
	if (!a && !b) return true;
	if (!a || !b) return false;
	return a.href === b.href && a.anchor === b.anchor && a.tooltip === b.tooltip;
}

function setHyperlinkTarget(
	container: XmlElement,
	link: HyperlinkInfo,
	allocator: RelationshipAllocator,
): void {
	container.removeAttributeNS(REL_NS, 'id');
	if (link.href)
		container.setAttributeNS(REL_NS, 'r:id', allocator.addExternalHyperlink(link.href));
	if (link.anchor) setAttribute(container, 'anchor', link.anchor);
	else container.removeAttributeNS(WORD_NS, 'anchor');
	if (link.tooltip) setAttribute(container, 'tooltip', link.tooltip);
	else container.removeAttributeNS(WORD_NS, 'tooltip');
	if (!getW(container, 'history')) setAttribute(container, 'history', '1');
}

/** Builds the paragraph's ordered top-level run/hyperlink nodes from the model's flat run list. */
export function buildRunNodes(
	doc: XmlDocument,
	runs: TextRun[],
	base: TextRun[] | undefined,
	oldSlots: RunSlot[],
	allocator: RelationshipAllocator,
): XmlElement[] {
	const nodes: XmlElement[] = [];
	let i = 0;
	while (i < runs.length) {
		const run = runs[i];
		if (!run.link) {
			nodes.push(createRun(doc, run, base?.[i], oldSlots[i]?.element, allocator));
			i++;
			continue;
		}
		const start = i;
		const linkValue = run.link;
		while (i < runs.length && sameLink(runs[i].link, linkValue)) i++;
		const group = runs.slice(start, i);
		const groupOld = oldSlots.slice(start, i);
		const baseGroup = base?.slice(start, i);
		const reusable =
			groupOld.length === group.length &&
			groupOld.every((slot): slot is Extract<RunSlot, { kind: 'link' }> => slot?.kind === 'link') &&
			new Set(groupOld.map((slot) => slot.container)).size === 1 &&
			baseGroup?.length === group.length &&
			baseGroup.every((run) => sameLink(run.link, linkValue));
		let container: XmlElement;
		if (reusable) {
			container = (groupOld[0] as Extract<RunSlot, { kind: 'link' }>).container;
			for (const child of Array.from(container.childNodes)) container.removeChild(child);
		} else {
			container = makeW(doc, 'hyperlink');
			setHyperlinkTarget(container, linkValue, allocator);
		}
		group.forEach((run, index) =>
			container.appendChild(
				createRun(doc, run, baseGroup?.[index], groupOld[index]?.element, allocator),
			),
		);
		nodes.push(container);
	}
	return nodes;
}
