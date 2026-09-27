// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import JSZip from 'jszip';
import type { DocumentModel, PendingMediaPart } from './model.js';
import { buildXml, parseXml, type XmlDocument } from './xml.js';
import type { PackageContext } from './parse.js';
import { applyModel } from './write.js';
import { applyNumberingCatalog } from './numbering-package.js';
import { applyHeaderFooterEdits, sectionLayout } from './write-header-footer.js';
import { applyTrackChangesSetting } from './settings.js';
import { applyComments } from './write-comments.js';
import {
	buildRelationshipsXml,
	ensureContentTypeDefault,
	parseRelationships,
} from './package-parts.js';

const RELS_PART = 'word/_rels/document.xml.rels';
const CONTENT_TYPES_PART = '[Content_Types].xml';
const extensionOf = (partName: string): string => partName.split('.').pop()?.toLowerCase() ?? '';

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
	if (
		binding &&
		JSON.stringify(sectionLayout(model.sections)) !==
			JSON.stringify(sectionLayout(binding.base.sections))
	)
		throw new Error(
			'Editing sections or page setup is not supported; source section properties are preserved unchanged.',
		);
	if (
		binding &&
		(JSON.stringify(model.footnotes) !== JSON.stringify(binding.base.footnotes) ||
			JSON.stringify(model.endnotes) !== JSON.stringify(binding.base.endnotes))
	)
		throw new Error(
			'Editing footnote or endnote text is not supported; source footnotes.xml/endnotes.xml are preserved unchanged.',
		);
	if (!binding && (model.sections?.length || model.footnotes?.length || model.endnotes?.length))
		throw new Error(
			'Sections, headers, footers, footnotes and endnotes are not supported by the standalone DOCX writer; they would be silently dropped.',
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
	if (newRelationships.length) {
		const relationships = parseRelationships(await zip.file(RELS_PART)?.async('string'));
		let contentTypesXml = await zip.file(CONTENT_TYPES_PART)?.async('string');
		for (const relationship of newRelationships) {
			relationships.set(relationship.id, {
				target: relationship.target,
				mode: relationship.mode,
				type: relationship.type,
			});
			if (relationship.partName) {
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
		}
		zip.file(RELS_PART, buildRelationshipsXml(relationships));
		if (contentTypesXml) zip.file(CONTENT_TYPES_PART, contentTypesXml);
	}
	if (binding) await applyHeaderFooterEdits(zip, model, binding.base);
	await applyNumberingCatalog(zip, model, binding);
	if (model.trackChanges !== binding?.base.trackChanges)
		await applyTrackChangesSetting(zip, model.trackChanges === true);
	if (JSON.stringify(model.comments ?? []) !== JSON.stringify(binding?.base.comments ?? []))
		await applyComments(zip, model.comments ?? []);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
