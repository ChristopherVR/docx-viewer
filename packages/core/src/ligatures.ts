import { elements, type XmlDocument, type XmlElement } from './xml.js';

export const LIGATURE_VALUES = [
	'none',
	'standard',
	'contextual',
	'historical',
	'discretional',
	'standardContextual',
	'standardHistorical',
	'contextualHistorical',
	'standardDiscretional',
	'contextualDiscretional',
	'historicalDiscretional',
	'standardContextualHistorical',
	'standardContextualDiscretional',
	'standardHistoricalDiscretional',
	'contextualHistoricalDiscretional',
	'all',
] as const;
export type Ligatures = (typeof LIGATURE_VALUES)[number];
export const isLigatures = (value: unknown): value is Ligatures =>
	typeof value === 'string' && (LIGATURE_VALUES as readonly string[]).includes(value);
export const WORD_2010_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XMLNS = 'http://www.w3.org/2000/xmlns/';
export const ligatureElements = (props: XmlElement) =>
	elements(props).filter(
		(element) => element.namespaceURI === WORD_2010_NS && element.localName === 'ligatures',
	);
export const parseLigatures = (props: XmlElement): Ligatures | undefined => {
	const value = ligatureElements(props)[0]?.getAttributeNS(WORD_2010_NS, 'val');
	return isLigatures(value) ? value : undefined;
};

/** Declare extensions as ignorable for pre-2010 consumers, preserving existing namespace bindings. */
function extensionPrefix(doc: XmlDocument, ns: string, preferred: string): string {
	const root = doc.documentElement;
	let prefix = preferred;
	for (
		let index = 1;
		root.lookupNamespaceURI(prefix) && root.lookupNamespaceURI(prefix) !== ns;
		index++
	)
		prefix = `${preferred}${index}`;
	root.setAttributeNS(XMLNS, `xmlns:${prefix}`, ns);
	return prefix;
}
export function writeLigatures(
	doc: XmlDocument,
	props: XmlElement,
	value: Ligatures | undefined,
): void {
	for (const element of ligatureElements(props)) props.removeChild(element);
	if (value === undefined) return;
	const prefix = extensionPrefix(doc, WORD_2010_NS, 'w14');
	const mc = extensionPrefix(doc, MC, 'mc');
	const ignored = new Set(
		(doc.documentElement.getAttributeNS(MC, 'Ignorable') ?? '').split(/\s+/).filter(Boolean),
	);
	ignored.add(prefix);
	doc.documentElement.setAttributeNS(MC, `${mc}:Ignorable`, [...ignored].join(' '));
	const element = doc.createElementNS(WORD_2010_NS, `${prefix}:ligatures`);
	element.setAttributeNS(WORD_2010_NS, `${prefix}:val`, value);
	props.appendChild(element);
}
