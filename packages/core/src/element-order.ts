// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word validates property children against their schema sequence; edits must keep that order.
import { isElement, WORD_NS, type XmlElement } from './xml.js';

/** `CT_RPr` child order (ECMA-376 §17.3.2.28). */
export const RPR_ORDER = [
	'rStyle',
	'rFonts',
	'b',
	'bCs',
	'i',
	'iCs',
	'caps',
	'smallCaps',
	'strike',
	'dstrike',
	'outline',
	'shadow',
	'emboss',
	'imprint',
	'noProof',
	'snapToGrid',
	'vanish',
	'webHidden',
	'color',
	'spacing',
	'w',
	'kern',
	'position',
	'sz',
	'szCs',
	'highlight',
	'u',
	'effect',
	'bdr',
	'shd',
	'fitText',
	'vertAlign',
	'rtl',
	'cs',
	'em',
	'lang',
	'eastAsianLayout',
	'specVanish',
	'oMath',
	'ins',
	'del',
	'moveFrom',
	'moveTo',
	'rPrChange',
];

/**
 * Sorts an element's children into `order`. Elements outside the list (extensions) keep their
 * position relative to the known element before them. Nothing moves when already in order.
 */
export function orderChildren(parent: XmlElement, order: readonly string[]): void {
	const elements = Array.from(parent.childNodes).filter(isElement);
	let previousRank = -1;
	const ranked = elements.map((element, index) => {
		const rank = element.namespaceURI === WORD_NS ? order.indexOf(element.localName ?? '') : -1;
		if (rank >= 0) previousRank = rank;
		return { element, index, rank: rank >= 0 ? rank : previousRank + 0.5 };
	});
	const sorted = [...ranked].sort((a, b) => a.rank - b.rank || a.index - b.index);
	if (sorted.every((entry, index) => entry.element === elements[index])) return;
	for (const { element } of sorted) parent.appendChild(element);
}
