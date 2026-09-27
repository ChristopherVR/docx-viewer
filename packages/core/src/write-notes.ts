// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes edited footnote/endnote text back into word/footnotes.xml and word/endnotes.xml, one
// note element at a time. Unedited notes and the separator notes stay byte-identical.
import type JSZip from 'jszip';
import type { DocumentModel, Note } from './model.js';
import { buildXml, children, getW, parseXml } from './xml.js';
import { applyBlocks } from './write.js';

const PARTS = {
	footnote: { part: 'word/footnotes.xml', key: 'footnotes' },
	endnote: { part: 'word/endnotes.xml', key: 'endnotes' },
} as const;

export async function applyNoteEdits(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel,
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
		if (!file) throw new Error(`Cannot edit ${kind}s: ${part} is missing from the package.`);
		const doc = parseXml(await file.async('string'));
		const elements = new Map(
			children(doc.documentElement, kind).map((element) => [getW(element, 'id'), element]),
		);
		for (const note of edited) {
			const element = elements.get(note.id);
			const original = baseNotes.get(note.id);
			if (!element || !original)
				throw new Error(
					`Adding new ${kind}s is not supported yet; ${kind} ${note.id} was not saved.`,
				);
			applyBlocks(doc, element, note.blocks, original.blocks, undefined, contentWidthTwips);
		}
		zip.file(part, buildXml(doc));
	}
}
