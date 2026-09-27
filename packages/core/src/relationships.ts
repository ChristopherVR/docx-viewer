// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { isElement, parseXml, type XmlElement } from './xml.js';

export const RELATIONSHIP_NS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface Relationship {
	id: string;
	type: string;
	target: string;
}

/** Parses a `_rels/*.rels` part into a map keyed by relationship id. */
export function parseRelationships(xml: string | undefined): Map<string, Relationship> {
	const map = new Map<string, Relationship>();
	if (!xml) return map;
	const document = parseXml(xml);
	for (const element of Array.from(document.documentElement.childNodes).filter(isElement)) {
		if (element.localName !== 'Relationship') continue;
		const id = element.getAttribute('Id');
		const type = element.getAttribute('Type');
		const target = element.getAttribute('Target');
		if (id && type && target) map.set(id, { id, type, target });
	}
	return map;
}

/** Resolves a relationship `Target` against the part that declared it (usually `word/...`). */
export function resolvePartPath(basePart: string, target: string): string {
	if (target.startsWith('/')) return target.slice(1);
	const baseDir = basePart.slice(0, basePart.lastIndexOf('/') + 1);
	const segments = `${baseDir}${target}`.split('/');
	const resolved: string[] = [];
	for (const segment of segments) {
		if (segment === '.' || segment === '') continue;
		if (segment === '..') resolved.pop();
		else resolved.push(segment);
	}
	return resolved.join('/');
}

/** Reads the `r:id` relationship reference attribute on an element, e.g. `w:headerReference`. */
export function getRelationshipId(element: XmlElement): string | undefined {
	return element.getAttributeNS(RELATIONSHIP_NS, 'id') || element.getAttribute('r:id') || undefined;
}
