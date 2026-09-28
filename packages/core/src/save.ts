// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import JSZip from 'jszip';
import type { DocumentModel, PendingMediaPart } from './model.js';
import { buildXml, parseXml, type XmlDocument } from './xml.js';
import type { PackageContext } from './parse.js';
import { applyModel } from './write.js';
import { applyNumberingCatalog } from './numbering-package.js';
import { applyHeaderFooterEdits } from './write-header-footer.js';
import { applyNoteEdits } from './write-notes.js';
import { writeNewRelationships } from './part-relationships.js';
import { applySettingsFlag, applyTrackChangesSetting } from './settings.js';
import { applyComments } from './write-comments.js';
import { numberCommentIds } from './comment-spans.js';
import { maxWordId, numberRevisionIds } from './revision-ids.js';
import { parseRelationships } from './package-parts.js';

const RELS_PART = 'word/_rels/document.xml.rels';

const contexts = new WeakMap<DocumentModel, { context: PackageContext; base: DocumentModel }>();

export function remember(
	model: DocumentModel,
	context: PackageContext,
	base: DocumentModel = context.base,
): void {
	contexts.set(model, { context, base: structuredClone(base) });
}
function newDocument(): XmlDocument {
	return parseXml(
		'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>',
	);
}

export async function saveDocx(
	model: DocumentModel,
	pendingMedia?: ReadonlyMap<string, PendingMediaPart>,
): Promise<Uint8Array> {
	const ids = new Set<string>();
	for (const block of model.blocks) {
		if (ids.has(block.id)) throw new Error(`Duplicate document block id: ${block.id}`);
		ids.add(block.id);
		if (block.type === 'table')
			for (const row of block.rows)
				for (const cell of row)
					for (const paragraph of cell.paragraphs) {
						if (ids.has(paragraph.id))
							throw new Error(`Duplicate document node id: ${paragraph.id}`);
						ids.add(paragraph.id);
					}
	}
	const binding = contexts.get(model);
	if (binding && JSON.stringify(model) === JSON.stringify(binding.base))
		return new Uint8Array(binding.context.original);
	model = numberRevisionIds(numberCommentIds(model), maxWordId(binding?.context.sourceXml));
	if (
		binding &&
		JSON.stringify(model.paragraphStyles) !== JSON.stringify(binding.base.paragraphStyles)
	)
		throw new Error(
			'Editing the paragraph style catalog is not supported; source styles.xml is preserved unchanged.',
		);
	if (!binding && model.paragraphStyles)
		throw new Error(
			'Creating or editing paragraph styles is not supported by the standalone DOCX writer.',
		);
	if (!binding && model.sections?.some((section) => section.headers || section.footers))
		throw new Error(
			'Headers and footers are not supported by the standalone DOCX writer; they would be silently dropped.',
		);
	if (
		binding &&
		JSON.stringify(model.characterStyles) !== JSON.stringify(binding.base.characterStyles)
	)
		throw new Error(
			'Editing the character style catalog is not supported; source styles.xml is preserved unchanged.',
		);
	if (binding && JSON.stringify(model.tableStyles) !== JSON.stringify(binding.base.tableStyles))
		throw new Error(
			'Editing the table style catalog is not supported; source styles.xml is preserved unchanged.',
		);
	if (binding && JSON.stringify(model.theme) !== JSON.stringify(binding.base.theme))
		throw new Error(
			'Editing the theme catalog is not supported; source theme1.xml/settings.xml are preserved unchanged.',
		);
	const zip = binding ? await JSZip.loadAsync(binding.context.original) : new JSZip();
	const document = binding ? parseXml(binding.context.sourceXml) : newDocument();
	const existingRelationships = parseRelationships(await zip.file(RELS_PART)?.async('string'));
	const { newRelationships } = applyModel(
		document,
		model,
		binding?.base.blocks ?? [],
		existingRelationships.keys(),
		binding?.base.sections ?? [],
	);
	zip.file('word/document.xml', buildXml(document));
	if (!binding) {
		zip.file(
			'[Content_Types].xml',
			'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
		);
		zip.file(
			'_rels/.rels',
			'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
		);
	}
	await writeNewRelationships(zip, 'word/document.xml', newRelationships, pendingMedia);
	if (binding) await applyHeaderFooterEdits(zip, model, binding.base, pendingMedia);
	// A new document starts without notes, so every note is created along with its part.
	await applyNoteEdits(
		zip,
		model,
		binding?.base ?? { ...model, footnotes: [], endnotes: [] },
		pendingMedia,
	);
	await applyNumberingCatalog(zip, model, binding);
	if (Boolean(model.evenAndOddHeaders) !== Boolean(binding?.base.evenAndOddHeaders))
		await applySettingsFlag(zip, 'evenAndOddHeaders', Boolean(model.evenAndOddHeaders));
	if (model.trackChanges !== binding?.base.trackChanges)
		await applyTrackChangesSetting(zip, model.trackChanges === true);
	if (JSON.stringify(model.comments ?? []) !== JSON.stringify(binding?.base.comments ?? []))
		await applyComments(zip, model.comments ?? []);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
