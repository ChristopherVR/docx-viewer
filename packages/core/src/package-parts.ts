// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Reads and writes the OPC plumbing (relationships, content types) that drawings and
// hyperlinks depend on. Kept separate from parse.ts/save.ts to stay under the module size limit.
import { isElement, parseXml, type XmlElement } from './xml.js';

export interface Relationship {
	target: string;
	mode: 'Internal' | 'External';
	type: string;
}

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

/** Parses `word/_rels/document.xml.rels`; returns an empty map when the part is absent. */
export function parseRelationships(xml: string | undefined): Map<string, Relationship> {
	const map = new Map<string, Relationship>();
	if (!xml) return map;
	const doc = parseXml(xml);
	const root = Array.from(doc.childNodes).filter(isElement)[0];
	if (!root) return map;
	for (const node of Array.from(root.childNodes).filter(isElement)) {
		if (node.localName !== 'Relationship') continue;
		const id = node.getAttribute('Id');
		const target = node.getAttribute('Target');
		if (!id || target == null) continue;
		map.set(id, {
			target,
			mode: node.getAttribute('TargetMode') === 'External' ? 'External' : 'Internal',
			type: node.getAttribute('Type') ?? '',
		});
	}
	return map;
}

/** Resolves a relationship target that is relative to `word/` (the only base part this codec writes into). */
export function resolveInternalTarget(target: string): string {
	if (target.startsWith('/')) return target.replace(/^\//, '');
	return `word/${target.replace(/^\.\//, '')}`;
}

/** Serializes relationships back into a `.rels` part, preserving `xml:id`-free minimal formatting. */
export function buildRelationshipsXml(relationships: Map<string, Relationship>): string {
	const items = Array.from(relationships.entries())
		.map(([id, rel]) => {
			const mode = rel.mode === 'External' ? ' TargetMode="External"' : '';
			return `<Relationship Id="${id}" Type="${rel.type}" Target="${escapeAttr(rel.target)}"${mode}/>`;
		})
		.join('');
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}">${items}</Relationships>`;
}

function escapeAttr(value: string): string {
	return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

export interface ContentTypes {
	defaults: Map<string, string>;
	overrides: Map<string, string>;
	raw?: XmlElement;
}

/** Parses `[Content_Types].xml`; returns empty maps when the part is absent. */
export function parseContentTypes(xml: string | undefined): ContentTypes {
	const defaults = new Map<string, string>();
	const overrides = new Map<string, string>();
	if (!xml) return { defaults, overrides };
	const doc = parseXml(xml);
	const root = Array.from(doc.childNodes).filter(isElement)[0];
	if (!root) return { defaults, overrides };
	for (const node of Array.from(root.childNodes).filter(isElement)) {
		if (node.localName === 'Default') {
			const extension = node.getAttribute('Extension');
			const contentType = node.getAttribute('ContentType');
			if (extension && contentType) defaults.set(extension.toLowerCase(), contentType);
		} else if (node.localName === 'Override') {
			const partName = node.getAttribute('PartName');
			const contentType = node.getAttribute('ContentType');
			if (partName && contentType) overrides.set(partName, contentType);
		}
	}
	return { defaults, overrides, raw: root };
}

export function contentTypeForPart(types: ContentTypes, partName: string): string | undefined {
	const overridePath = partName.startsWith('/') ? partName : `/${partName}`;
	if (types.overrides.has(overridePath)) return types.overrides.get(overridePath);
	const extension = partName.split('.').pop()?.toLowerCase();
	return extension ? types.defaults.get(extension) : undefined;
}

/** Adds a `Default` extension mapping if one is not already present, returning the (possibly unchanged) XML. */
export function ensureContentTypeDefault(
	xml: string | undefined,
	extension: string,
	contentType: string,
): string {
	const types = parseContentTypes(xml);
	if (types.defaults.has(extension.toLowerCase())) return xml ?? buildContentTypesXml(types);
	types.defaults.set(extension.toLowerCase(), contentType);
	return buildContentTypesXml(types);
}

function buildContentTypesXml(types: ContentTypes): string {
	const defaults = Array.from(types.defaults.entries())
		.map(
			([extension, contentType]) =>
				`<Default Extension="${extension}" ContentType="${contentType}"/>`,
		)
		.join('');
	const overrides = Array.from(types.overrides.entries())
		.map(
			([partName, contentType]) =>
				`<Override PartName="${escapeAttr(partName)}" ContentType="${contentType}"/>`,
		)
		.join('');
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CONTENT_TYPES_NS}">${defaults}${overrides}</Types>`;
}
