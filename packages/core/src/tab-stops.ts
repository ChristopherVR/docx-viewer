// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Paragraph tab stops (`w:pPr/w:tabs`) and the canonical `w:pPr` child order Word requires.
import type { Paragraph, TabStop } from './model.js';
import { orderChildren } from './element-order.js';
import { isStTabJc, isStTabTlc } from './generated/wml-simple-types.js';
import { enumValue } from './parse-diagnostics.js';
import { parseSignedTwips } from './simple-types.js';
import {
	children,
	getW,
	isElement,
	makeW,
	WORD_NS,
	type XmlDocument,
	type XmlElement,
} from './xml.js';

export function parseTabStops(tabs: XmlElement | undefined): TabStop[] {
	const stops: TabStop[] = [];
	for (const tab of tabs ? children(tabs, 'tab') : []) {
		const align = enumValue(isStTabJc, getW(tab, 'val'), 'w:tab/@w:val');
		const pos = parseSignedTwips(getW(tab, 'pos'));
		if (!align || pos === undefined) continue;
		const leader = enumValue(isStTabTlc, getW(tab, 'leader'), 'w:tab/@w:leader');
		stops.push({ posTwips: pos, align, ...(leader ? { leader } : {}) });
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

/** Sorts `w:pPr` children into schema order. */
export function orderParagraphProperties(props: XmlElement): void {
	orderChildren(props, PPR_ORDER);
}
