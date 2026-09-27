import { getW, WORD_NS, type XmlElement } from './xml.js';

/** Only ordinary text-wrapping breaks are represented by a model newline. */
export function hasSpecialBreak(element: XmlElement): boolean {
	for (const attribute of Array.from(element.attributes)) {
		if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
		if (attribute.namespaceURI !== WORD_NS) return true;
		if (attribute.localName === 'type') {
			if (attribute.value !== 'textWrapping') return true;
		} else if (attribute.localName === 'clear') {
			if (attribute.value !== 'none') return true;
		} else return true;
	}
	return false;
}

/** `w:br/@w:type`; absent type defaults to `textWrapping` per the OOXML schema. */
export function classifyBreak(element: XmlElement): 'page' | 'column' | 'textWrapping' {
	const type = getW(element, 'type');
	return type === 'page' || type === 'column' ? type : 'textWrapping';
}

/**
 * A page or column break with no other attributes. These are modeled explicitly (`TextRun.break`)
 * and are safe to relocate on edit; other special breaks (e.g. `w:clear`) remain unsupported.
 */
export function isModeledBreak(element: XmlElement): boolean {
	const kind = classifyBreak(element);
	if (kind === 'textWrapping') return false;
	for (const attribute of Array.from(element.attributes)) {
		if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
		if (attribute.namespaceURI !== WORD_NS || attribute.localName !== 'type') return false;
	}
	return true;
}
