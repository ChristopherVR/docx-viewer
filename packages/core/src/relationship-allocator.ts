// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Assigns fresh relationship ids when write.ts introduces a new external hyperlink or a
// newly inserted picture, and records what save.ts must add to word/_rels/document.xml.rels.
import { DocPrIdAllocator } from './docpr-ids.js';
import { isElement, REL_NS, type XmlDocument, type XmlElement } from './xml.js';

export const HYPERLINK_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
export const IMAGE_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const SAFE_HYPERLINK_SCHEMES = new Set(['http:', 'https:', 'mailto:']);

/** Only http(s) and mailto targets are accepted; anything else (notably `javascript:`) is rejected. */
export function isSafeHyperlinkHref(href: string): boolean {
	const match = /^([a-z][a-z0-9+.-]*:)/i.exec(href.trim());
	const scheme = match?.[1];
	return scheme !== undefined && SAFE_HYPERLINK_SCHEMES.has(scheme.toLowerCase());
}

export interface NewRelationship {
	id: string;
	type: string;
	target: string;
	mode: 'Internal' | 'External';
	/** Absolute package part name (e.g. `word/media/image1.png`), set only for newly added media parts. */
	partName?: string;
}

/** Walks every element/attribute in a parsed document.xml to find relationship ids already in use. */
export function scanUsedRelationshipIds(doc: XmlDocument): Set<string> {
	const used = new Set<string>();
	const visit = (node: XmlElement) => {
		for (const attribute of Array.from(node.attributes))
			if (attribute.namespaceURI === REL_NS) used.add(attribute.value);
		for (const child of Array.from(node.childNodes)) if (isElement(child)) visit(child);
	};
	if (doc.documentElement) visit(doc.documentElement);
	return used;
}

/** Allocates unused `rIdN` ids for new relationships introduced while applying a model diff. */
export class RelationshipAllocator {
	private readonly used: Set<string>;
	private readonly created: NewRelationship[] = [];
	private readonly imagePartToRelId = new Map<string, string>();

	/** Shared drawing-id allocator, so `wp:docPr` ids stay unique across every part of a save. */
	readonly docPrIds: DocPrIdAllocator;

	constructor(existingIds: Iterable<string>, docPrIds: DocPrIdAllocator = new DocPrIdAllocator()) {
		this.used = new Set(existingIds);
		this.docPrIds = docPrIds;
	}

	private nextId(): string {
		let index = 1;
		while (this.used.has(`rId${index}`)) index++;
		const id = `rId${index}`;
		this.used.add(id);
		return id;
	}

	/** Registers a brand-new external hyperlink target and returns its relationship id. */
	addExternalHyperlink(href: string): string {
		if (!isSafeHyperlinkHref(href))
			throw new Error(`Unsupported hyperlink target: only http(s) and mailto links are supported.`);
		const id = this.nextId();
		this.created.push({ id, type: HYPERLINK_RELATIONSHIP_TYPE, target: href, mode: 'External' });
		return id;
	}

	/** Registers (once per part) a newly inserted image part and returns its relationship id. */
	addImage(partName: string): string {
		const existing = this.imagePartToRelId.get(partName);
		if (existing) return existing;
		const id = this.nextId();
		this.created.push({
			id,
			type: IMAGE_RELATIONSHIP_TYPE,
			target: partName.replace(/^word\//, ''),
			mode: 'Internal',
			partName,
		});
		this.imagePartToRelId.set(partName, id);
		return id;
	}

	get newRelationships(): readonly NewRelationship[] {
		return this.created;
	}
}
