// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes edited footnote/endnote text back into word/footnotes.xml and word/endnotes.xml, one
// note element at a time. Unedited notes and the separator notes stay byte-identical.
import type JSZip from 'jszip';
import type { DocumentModel, Note, PendingMediaPart } from './model.js';
import { allocatorForPart, writeNewRelationships } from './part-relationships.js';
import { buildXml, children, getW, parseXml } from './xml.js';
import { applyBlocks } from './write.js';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts.js';

const PARTS = {
	footnote: { part: 'word/footnotes.xml', key: 'footnotes' },
	endnote: { part: 'word/endnotes.xml', key: 'endnotes' },
} as const;
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A notes part holding only Word's separator and continuation separator notes. */
function emptyNotesPart(kind: keyof typeof PARTS): string {
	const tag = `w:${kind}`;
	const separator = (type: string, id: string, mark: string) =>
		`<${tag} w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${mark}/></w:r></w:p></${tag}>`;
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s xmlns:w="${WORD_NS}">${separator('separator', '-1', 'separator')}${separator('continuationSeparator', '0', 'continuationSeparator')}</w:${kind}s>`;
}

/** Creates the notes part, its relationship and content type the first time a note is added. */
async function createNotesPart(zip: JSZip, kind: keyof typeof PARTS): Promise<string> {
	await ensureContentTypeOverride(
		zip,
		PARTS[kind].part,
		`application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml`,
	);
	await ensureDocumentRelationship(zip, `${RELATIONSHIPS}/${kind}s`, `${kind}s.xml`);
	return emptyNotesPart(kind);
}

export async function applyNoteEdits(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel,
	pendingMedia?: ReadonlyMap<string, PendingMediaPart>,
): Promise<void> {
	const contentWidthTwips = Math.round(
		(model.page.width - model.page.marginLeft - model.page.marginRight) * 15,
	);
	for (const [kind, { part, key }] of Object.entries(PARTS) as [
		keyof typeof PARTS,
		(typeof PARTS)[keyof typeof PARTS],
	][]) {
		const baseNotes = new Map<string, Note>((base[key] ?? []).map((note) => [note.id, note]));
		const edited = (model[key] ?? []).filter((note) => {
			const original = baseNotes.get(note.id);
			return !original || JSON.stringify(original.blocks) !== JSON.stringify(note.blocks);
		});
		if (!edited.length) continue;
		const file = zip.file(part);
		const doc = parseXml(file ? await file.async('string') : await createNotesPart(zip, kind));
		const elements = new Map(
			children(doc.documentElement, kind).map((element) => [getW(element, 'id'), element]),
		);
		const allocator = await allocatorForPart(zip, part, doc);
		for (const note of edited) {
			let element = elements.get(note.id);
			const original = baseNotes.get(note.id);
			if (original && !element)
				throw new Error(`${kind} ${note.id} is missing from ${part}; it was not saved.`);
			if (!element) {
				element = doc.createElementNS(WORD_NS, `w:${kind}`);
				element.setAttributeNS(WORD_NS, 'w:id', note.id);
				doc.documentElement.appendChild(element);
			}
			applyBlocks(doc, element, note.blocks, original?.blocks ?? [], allocator, contentWidthTwips);
		}
		zip.file(part, buildXml(doc));
		await writeNewRelationships(zip, part, allocator.newRelationships, pendingMedia);
	}
}
