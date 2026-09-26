import { WORD_NS, type XmlElement } from './xml.js';

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
