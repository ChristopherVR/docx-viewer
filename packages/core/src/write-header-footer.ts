// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes edited header and footer content back into its own package part. Only parts whose blocks
// changed are rewritten; everything else in the package stays byte-identical.
import type JSZip from 'jszip';
import type {
	Block,
	DocumentModel,
	HeaderFooterSlots,
	PendingMediaPart,
	SectionProperties,
} from './model.js';
import {
	allocatorForPart,
	relationshipsPartFor,
	writeNewRelationships,
} from './part-relationships.js';
import type { DocPrIdAllocator } from './docpr-ids.js';
import { buildXml, parseXml } from './xml.js';
import { applyBlocks } from './write.js';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts.js';

type Kind = 'headers' | 'footers';
const KINDS: Kind[] = ['headers', 'footers'];
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PART_KINDS = {
	headers: { root: 'hdr', relation: 'header', contentType: 'header' },
	footers: { root: 'ftr', relation: 'footer', contentType: 'footer' },
} as const;

/** Header/footer blocks by part name. Throws when two references to one part disagree. */
function partContents(sections: SectionProperties[] | undefined): Map<string, Block[]> {
	const parts = new Map<string, Block[]>();
	for (const section of sections ?? [])
		for (const kind of KINDS)
			for (const content of Object.values(
				section[kind] ?? {},
			) as HeaderFooterSlots[keyof HeaderFooterSlots][]) {
				if (!content?.partName) continue;
				const existing = parts.get(content.partName);
				if (existing && JSON.stringify(existing) !== JSON.stringify(content.blocks))
					throw new Error(
						`Header/footer part ${content.partName} is shared by several sections; edit every use of it the same way.`,
					);
				parts.set(content.partName, content.blocks);
			}
	return parts;
}

/** Header/footer parts the model references that the loaded package does not contain yet. */
export function newHeaderFooterParts(
	model: DocumentModel,
	base: DocumentModel | undefined,
): Array<{ partName: string; kind: Kind }> {
	const existing = partContents(base?.sections);
	const found = new Map<string, Kind>();
	for (const section of model.sections ?? [])
		for (const kind of KINDS)
			for (const content of Object.values(section[kind] ?? {}) as Array<
				HeaderFooterSlots[keyof HeaderFooterSlots]
			>)
				if (content?.partName && !existing.has(content.partName)) found.set(content.partName, kind);
	return [...found].map(([partName, kind]) => ({ partName, kind }));
}

/**
 * Creates every header/footer part the model added: an empty part, its content-type override and
 * its document relationship. Returns the relationship id of each new part so `w:sectPr` can
 * reference it. The part's blocks are written afterwards by `applyHeaderFooterEdits`.
 */
export async function createHeaderFooterParts(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel | undefined,
	docPrIds?: DocPrIdAllocator,
): Promise<Map<string, string>> {
	const ids = new Map<string, string>();
	for (const { partName, kind } of newHeaderFooterParts(model, base)) {
		const info = PART_KINDS[kind];
		if (!/^word\/[A-Za-z0-9_-]+\.xml$/.test(partName))
			throw new Error(`Unsupported ${info.relation} part name: ${partName}`);
		if (zip.file(partName))
			throw new Error(
				`Cannot add ${info.relation} part ${partName}: the package already contains a part with that name that no section references.`,
			);
		zip.file(
			partName,
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${info.root} xmlns:w="${WORD_NS}" xmlns:r="${RELATIONSHIPS}"/>`,
		);
		const content = (model.sections ?? [])
			.flatMap((section) => Object.values(section[kind] ?? {}))
			.find((content) => content?.partName === partName);
		const source = content?.sourcePartName && zip.file(content.sourcePartName);
		if (source) {
			const xml = await source.async('string');
			zip.file(
				partName,
				docPrIds
					? xml.replace(
							/(<(?:[\w.-]+:)?docPr\b[^>]*?\sid\s*=\s*["'])\d+(["'])/g,
						(_match: string, start: string, end: string) => `${start}${docPrIds.next()}${end}`,
						)
					: xml,
			);
			const rels = zip.file(relationshipsPartFor(content!.sourcePartName!));
			if (rels) zip.file(relationshipsPartFor(partName), await rels.async('uint8array'));
		}
		await ensureContentTypeOverride(
			zip,
			partName,
			`application/vnd.openxmlformats-officedocument.wordprocessingml.${info.contentType}+xml`,
		);
		ids.set(
			partName,
			await ensureDocumentRelationship(
				zip,
				`${RELATIONSHIPS}/${info.relation}`,
				partName.slice('word/'.length),
			),
		);
	}
	return ids;
}

/** Sections with header/footer content reduced to part references, for comparing layout only. */
export function sectionLayout(sections: SectionProperties[] | undefined): unknown {
	return (sections ?? []).map((section) => {
		const copy: Record<string, unknown> = { ...section };
		for (const kind of KINDS) {
			const slots = section[kind];
			if (slots)
				copy[kind] = Object.fromEntries(
					Object.entries(slots).map(([slot, content]) => [slot, content?.partName ?? null]),
				);
		}
		return copy;
	});
}

/** Rewrites every header/footer part whose blocks differ from the loaded baseline. */
export async function applyHeaderFooterEdits(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel,
	pendingMedia?: ReadonlyMap<string, PendingMediaPart>,
	docPrIds?: DocPrIdAllocator,
): Promise<void> {
	const next = partContents(model.sections);
	const previous = partContents(base.sections);
	const sources = new Map(
		(model.sections ?? []).flatMap((section) =>
			KINDS.flatMap((kind) =>
				Object.values(section[kind] ?? {}).flatMap((content) =>
					content?.partName && content.sourcePartName
						? [[content.partName, content.sourcePartName] as const]
						: [],
				),
			),
		),
	);
	const contentWidthTwips = Math.round(
		(model.page.width - model.page.marginLeft - model.page.marginRight) * 15,
	);
	const created = new Set(newHeaderFooterParts(model, base).map((part) => part.partName));
	for (const [partName, blocks] of next) {
		const original =
			previous.get(partName) ??
			previous.get(sources.get(partName) ?? '') ??
			(created.has(partName) ? [] : undefined);
		if (!original || JSON.stringify(original) === JSON.stringify(blocks)) continue;
		const file = zip.file(partName);
		if (!file) throw new Error(`Header/footer part ${partName} is missing from the package.`);
		const doc = parseXml(await file.async('string'));
		const allocator = await allocatorForPart(zip, partName, doc, docPrIds);
		applyBlocks(doc, doc.documentElement, blocks, original, allocator, contentWidthTwips);
		zip.file(partName, buildXml(doc));
		await writeNewRelationships(zip, partName, allocator.newRelationships, pendingMedia);
	}
}
