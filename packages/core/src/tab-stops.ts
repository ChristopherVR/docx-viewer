// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Paragraph tab stops (`w:pPr/w:tabs`) and the canonical `w:pPr` child order Word requires.
import type { Paragraph, TabStop } from './model.js';
import {
	children,
	getW,
	isElement,
	makeW,
	WORD_NS,
	type XmlDocument,
	type XmlElement,
} from './xml.js';

const ALIGNS = new Set<TabStop['align']>([
	'left',
	'center',
	'right',
	'decimal',
	'bar',
	'clear',
	'start',
	'end',
	'num',
]);
const LEADERS = new Set<NonNullable<TabStop['leader']>>([
	'none',
	'dot',
	'hyphen',
	'underscore',
	'heavy',
	'middleDot',
]);

export function parseTabStops(tabs: XmlElement | undefined): TabStop[] {
	const stops: TabStop[] = [];
	for (const tab of tabs ? children(tabs, 'tab') : []) {
		const align = getW(tab, 'val') as TabStop['align'] | undefined;
		const pos = Number(getW(tab, 'pos'));
		if (!align || !ALIGNS.has(align) || !Number.isSafeInteger(pos)) continue;
		const leader = getW(tab, 'leader') as TabStop['leader'];
		stops.push({ posTwips: pos, align, ...(leader && LEADERS.has(leader) ? { leader } : {}) });
	}
	return stops;
}

/** Rewrites `w:tabs` when the paragraph's tab stops changed. */
export function writeTabStops(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	if (base && JSON.stringify(paragraph.tabStops ?? []) === JSON.stringify(base.tabStops ?? []))
		return;
	for (const tabs of children(props, 'tabs')) props.removeChild(tabs);
	if (!paragraph.tabStops?.length) return;
	const tabs = makeW(doc, 'tabs');
	for (const stop of paragraph.tabStops) {
		const tab = makeW(doc, 'tab');
		tab.setAttributeNS(WORD_NS, 'w:val', stop.align);
		if (stop.leader) tab.setAttributeNS(WORD_NS, 'w:leader', stop.leader);
		tab.setAttributeNS(WORD_NS, 'w:pos', String(stop.posTwips));
		tabs.appendChild(tab);
	}
	props.appendChild(tabs);
}

/** `CT_PPr` child order (ECMA-376 §17.3.1.26); Word rejects paragraph properties out of order. */
const PPR_ORDER = [
	'pStyle',
	'keepNext',
	'keepLines',
	'pageBreakBefore',
	'framePr',
	'widowControl',
	'numPr',
	'suppressLineNumbers',
	'pBdr',
	'shd',
	'tabs',
	'suppressAutoHyphens',
	'kinsoku',
	'wordWrap',
	'overflowPunct',
	'topLinePunct',
	'autoSpaceDE',
	'autoSpaceDN',
	'bidi',
	'adjustRightInd',
	'snapToGrid',
	'spacing',
	'ind',
	'contextualSpacing',
	'mirrorIndents',
	'suppressOverlap',
	'jc',
	'textDirection',
	'textAlignment',
	'textboxTightWrap',
	'outlineLvl',
	'divId',
	'cnfStyle',
	'rPr',
	'sectPr',
	'pPrChange',
];

/**
 * Sorts `w:pPr` children into schema order. Elements outside the schema list (extensions) keep
 * their position relative to the known element before them.
 */
export function orderParagraphProperties(props: XmlElement): void {
	const elements = Array.from(props.childNodes).filter(isElement);
	let previousRank = -1;
	const ranked = elements.map((element, index) => {
		const rank = element.namespaceURI === WORD_NS ? PPR_ORDER.indexOf(element.localName ?? '') : -1;
		if (rank >= 0) previousRank = rank;
		return { element, index, rank: rank >= 0 ? rank : previousRank + 0.5 };
	});
	const sorted = [...ranked].sort((a, b) => a.rank - b.rank || a.index - b.index);
	if (sorted.every((entry, index) => entry.element === elements[index])) return;
	for (const { element } of sorted) props.appendChild(element);
}
