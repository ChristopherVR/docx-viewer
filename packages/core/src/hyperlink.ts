// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Resolves w:hyperlink targets (relationship-based external links, or w:anchor internal links)
// and simple, non-field-code HYPERLINK fields (w:fldSimple).
import type { HyperlinkInfo } from './model.js';
import type { Relationship } from './package-parts.js';
import { getR, getW, type XmlElement } from './xml.js';

/** Resolves a `w:hyperlink` element's target using the document's relationship map. */
export function resolveHyperlink(
	element: XmlElement,
	rels: ReadonlyMap<string, Relationship>,
): HyperlinkInfo {
	const info: HyperlinkInfo = {};
	const relId = getR(element, 'id');
	if (relId) {
		const rel = rels.get(relId);
		if (rel && rel.mode === 'External') info.href = rel.target;
	}
	const anchor = getW(element, 'anchor');
	if (anchor) info.anchor = anchor;
	const tooltip = getW(element, 'tooltip');
	if (tooltip) info.tooltip = tooltip;
	return info;
}

/** Parses a literal `HYPERLINK "target"` (optionally `\l anchor`) instruction from `w:fldSimple/@w:instr`. */
export function parseSimpleHyperlinkField(instr: string): HyperlinkInfo | undefined {
	const match = /^\s*HYPERLINK\s+"([^"]*)"(?:\s+\\l\s+"([^"]*)")?/i.exec(instr);
	if (!match) return undefined;
	const [, target, anchor] = match;
	if (anchor) return { anchor };
	if (target) return { href: target };
	return undefined;
}
