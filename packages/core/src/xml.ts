// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word-specific view of the shared XML model in `@christophervr/ooxml-xml`: the same parser and
// DOM, with the WordprocessingML namespace as the default for the element/attribute helpers.
import {
	NS,
	attr,
	buildXml,
	children as nsChildren,
	elements,
	first as nsFirst,
	isElement,
	makeNS,
	named as nsNamed,
	parseXml as parseOoxml,
	relAttr,
	textContent,
	type XmlDocument,
	type XmlElement,
	type XmlNode,
} from '@christophervr/ooxml-xml';

export const WORD_NS = NS.w;
/** OOXML relationships namespace used for `r:id`/`r:embed` attributes on drawings and hyperlinks. */
export const REL_NS = NS.r;
export type { XmlDocument, XmlElement, XmlNode };

export const parseXml = (xml: string): XmlDocument => parseOoxml(xml, { label: 'DOCX' });
export { buildXml, elements, isElement, makeNS, textContent };

export const named = (element: XmlElement | null | undefined, local: string): boolean =>
	nsNamed(element, local, WORD_NS);
export const children = (parent: ParentNode, local: string): XmlElement[] =>
	nsChildren(parent, local, WORD_NS);
export const first = (
	parent: ParentNode | null | undefined,
	local: string,
): XmlElement | undefined => nsFirst(parent, local, WORD_NS);
export const getW = (element: XmlElement | null | undefined, local: string): string | undefined =>
	attr(element, local, WORD_NS, 'w');
export const getR = (element: XmlElement | null | undefined, local: string): string | undefined =>
	relAttr(element, local);
export const makeW = (doc: XmlDocument, local: string): XmlElement =>
	doc.createElementNS(WORD_NS, `w:${local}`);
