// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

export const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
/** OOXML relationships namespace used for `r:id`/`r:embed` attributes on drawings and hyperlinks. */
export const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export type XmlDocument = Document;
export type XmlElement = Element;
export type XmlNode = Node;
type XmldomNode = Parameters<InstanceType<typeof XMLSerializer>['serializeToString']>[0];

export function parseXml(xml: string): XmlDocument {
	if (/<!DOCTYPE|<!ENTITY/i.test(xml))
		throw new Error('DOCX XML with DTD or entity declarations is not supported');
	let parseError = '';
	let parsed: XmlDocument;
	try {
		parsed = new DOMParser({
			onError: (level, message) => {
				if (level !== 'warning') parseError = message;
			},
		}).parseFromString(xml, 'application/xml') as unknown as XmlDocument;
	} catch (error) {
		throw new Error(`Invalid DOCX XML: ${error instanceof Error ? error.message : String(error)}`);
	}
	if (parseError || !parsed.documentElement)
		throw new Error(`Invalid DOCX XML: ${parseError || 'missing root element'}`);
	return parsed as unknown as XmlDocument;
}

export const buildXml = (document: XmlDocument): string =>
	new XMLSerializer().serializeToString(document as unknown as XmldomNode);
export const isElement = (node: Node): node is XmlElement => node.nodeType === 1;
export const elements = (parent: ParentNode): XmlElement[] =>
	Array.from(parent.childNodes).filter(isElement);
export const named = (element: XmlElement | null | undefined, local: string): boolean =>
	Boolean(
		element &&
		element.localName === local &&
		(!element.namespaceURI || element.namespaceURI === WORD_NS),
	);
export const children = (parent: ParentNode, local: string): XmlElement[] =>
	elements(parent).filter((element) => named(element, local));
export const first = (
	parent: ParentNode | null | undefined,
	local: string,
): XmlElement | undefined => (parent ? children(parent, local)[0] : undefined);
export const getW = (element: XmlElement | null | undefined, local: string): string | undefined =>
	element?.getAttributeNS(WORD_NS, local) ?? element?.getAttribute(`w:${local}`) ?? undefined;
export const getR = (element: XmlElement | null | undefined, local: string): string | undefined =>
	element?.getAttributeNS(REL_NS, local) ?? element?.getAttribute(`r:${local}`) ?? undefined;
export const makeW = (doc: XmlDocument, local: string): XmlElement =>
	doc.createElementNS(WORD_NS, `w:${local}`);
export const makeNS = (doc: XmlDocument, ns: string, qualified: string): XmlElement =>
	doc.createElementNS(ns, qualified);
export const textContent = (element: XmlElement): string => element.textContent ?? '';
