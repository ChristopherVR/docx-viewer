// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// OPC plumbing (relationships, content types) lives in the shared `@christophervr/ooxml-opc`; this
// module re-exports it and keeps the Word-specific base-part resolution.
export {
	buildRelationshipsXml,
	contentTypeForPart,
	ensureContentTypeDefault,
	parseContentTypes,
	parseRelationships,
	type ContentTypes,
	type Relationship,
} from '@christophervr/ooxml-opc';

/** Resolves a relationship target that is relative to `word/` (the only base part this codec writes into). */
export function resolveInternalTarget(target: string): string {
	if (target.startsWith('/')) return target.replace(/^\//, '');
	return `word/${target.replace(/^\.\//, '')}`;
}
