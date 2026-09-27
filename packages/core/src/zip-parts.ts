// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Minimal string-level surgery for OPC package plumbing (content types + relationships),
// used only to create parts (settings.xml, comments.xml) that a source package may lack.
import type JSZip from 'jszip';

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const DOCUMENT_RELS_PATH = 'word/_rels/document.xml.rels';

function attr(tag: string, name: string): string | undefined {
	return new RegExp(`${name}="([^"]*)"`).exec(tag)?.[1];
}

/** Adds an `<Override>` for the given part if [Content_Types].xml doesn't already declare one. */
export async function ensureContentTypeOverride(
	zip: JSZip,
	partName: string,
	contentType: string,
): Promise<void> {
	const file = zip.file(CONTENT_TYPES_PATH);
	const xml =
		(await file?.async('string')) ??
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
	const target = `/${partName}`;
	for (const tag of xml.match(/<Override\b[^>]*\/>/g) ?? [])
		if (attr(tag, 'PartName') === target) return;
	const insertion = `<Override PartName="${target}" ContentType="${contentType}"/>`;
	zip.file(CONTENT_TYPES_PATH, xml.replace('</Types>', `${insertion}</Types>`));
}

/** Returns the relationship id for `target`, creating the relationship (and rels part) if needed. */
export async function ensureDocumentRelationship(
	zip: JSZip,
	type: string,
	target: string,
): Promise<string> {
	const file = zip.file(DOCUMENT_RELS_PATH);
	const xml =
		(await file?.async('string')) ??
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
	let maxId = 0;
	for (const tag of xml.match(/<Relationship\b[^>]*\/>/g) ?? []) {
		const id = attr(tag, 'Id');
		const numeric = Number(id?.replace(/^rId/, ''));
		if (Number.isFinite(numeric)) maxId = Math.max(maxId, numeric);
		if (attr(tag, 'Type') === type && attr(tag, 'Target') === target && id) return id;
	}
	const id = `rId${maxId + 1}`;
	const insertion = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
	zip.file(DOCUMENT_RELS_PATH, xml.replace('</Relationships>', `${insertion}</Relationships>`));
	return id;
}
