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
import { allocatorForPart, writeNewRelationships } from './part-relationships.js';
import { buildXml, parseXml } from './xml.js';
import { applyBlocks } from './write.js';

type Kind = 'headers' | 'footers';
const KINDS: Kind[] = ['headers', 'footers'];

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
): Promise<void> {
	const next = partContents(model.sections);
	const previous = partContents(base.sections);
	const contentWidthTwips = Math.round(
		(model.page.width - model.page.marginLeft - model.page.marginRight) * 15,
	);
	for (const [partName, blocks] of next) {
		const original = previous.get(partName);
		if (!original || JSON.stringify(original) === JSON.stringify(blocks)) continue;
		const file = zip.file(partName);
		if (!file) throw new Error(`Header/footer part ${partName} is missing from the package.`);
		const doc = parseXml(await file.async('string'));
		const allocator = await allocatorForPart(zip, partName, doc);
		applyBlocks(doc, doc.documentElement, blocks, original, allocator, contentWidthTwips);
		zip.file(partName, buildXml(doc));
		await writeNewRelationships(zip, partName, allocator.newRelationships, pendingMedia);
	}
}
