// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes relationships added while applying edits to one package part (the document, a header,
// a footer or a notes part), plus the bytes and content types of newly inserted media.
import type JSZip from 'jszip';
import type { PendingMediaPart } from './model.js';
import {
	buildRelationshipsXml,
	ensureContentTypeDefault,
	parseRelationships,
} from './package-parts.js';
import {
	RelationshipAllocator,
	scanUsedRelationshipIds,
	type NewRelationship,
} from './relationship-allocator.js';
import type { XmlDocument } from './xml.js';

const CONTENT_TYPES_PART = '[Content_Types].xml';
const extensionOf = (partName: string): string => partName.split('.').pop()?.toLowerCase() ?? '';

/** The relationships part for `partName`, e.g. word/_rels/header1.xml.rels. */
export function relationshipsPartFor(partName: string): string {
	const slash = partName.lastIndexOf('/');
	return `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`;
}

/** An allocator that avoids every id already used by the part's XML or declared in its .rels. */
export async function allocatorForPart(
	zip: JSZip,
	partName: string,
	doc: XmlDocument,
): Promise<RelationshipAllocator> {
	const declared = parseRelationships(
		await zip.file(relationshipsPartFor(partName))?.async('string'),
	);
	return new RelationshipAllocator([...scanUsedRelationshipIds(doc), ...declared.keys()]);
}

/** Adds `created` relationships to the part's .rels and writes any newly inserted media. */
export async function writeNewRelationships(
	zip: JSZip,
	partName: string,
	created: readonly NewRelationship[],
	pendingMedia?: ReadonlyMap<string, PendingMediaPart>,
): Promise<void> {
	if (!created.length) return;
	const relsPath = relationshipsPartFor(partName);
	const relationships = parseRelationships(await zip.file(relsPath)?.async('string'));
	let contentTypesXml = await zip.file(CONTENT_TYPES_PART)?.async('string');
	for (const relationship of created) {
		relationships.set(relationship.id, {
			target: relationship.target,
			mode: relationship.mode,
			type: relationship.type,
		});
		if (!relationship.partName) continue;
		const pending = pendingMedia?.get(relationship.partName);
		if (!pending)
			throw new Error(`Missing bytes for newly inserted media part: ${relationship.partName}`);
		zip.file(relationship.partName, pending.bytes);
		contentTypesXml = ensureContentTypeDefault(
			contentTypesXml,
			extensionOf(relationship.partName),
			pending.contentType,
		);
	}
	zip.file(relsPath, buildRelationshipsXml(relationships));
	if (contentTypesXml) zip.file(CONTENT_TYPES_PART, contentTypesXml);
}
