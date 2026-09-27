// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Orchestrates section, header/footer and footnote/endnote parsing on top of block-parser.ts,
// sections.ts and notes.ts. Kept out of parse.ts so the hub file only needs one hook call.
import type JSZip from 'jszip';
import type { Block, HeaderFooterSlots, Note, SectionProperties } from './model.js';
import { first, getW, parseXml, type XmlElement } from './xml.js';
import { parseRelationships, resolvePartPath, type Relationship } from './relationships.js';
import { parseRawSections, type RawHeaderFooterRef } from './sections.js';
import { parseNotesPart } from './notes.js';
import { parseBlocksFromContainer } from './block-parser.js';

const FIELD_MARKUP = /<w:(?:fldSimple|instrText|fldChar)\b/;

async function readPart(zip: JSZip, path: string): Promise<string | undefined> {
	return zip.file(path)?.async('string');
}

interface SettingsInfo {
	evenAndOddHeaders: boolean;
	footnoteNumFmt?: string;
	endnoteNumFmt?: string;
}
function parseSettings(xml: string | undefined): SettingsInfo {
	if (!xml) return { evenAndOddHeaders: false };
	const root = parseXml(xml).documentElement;
	const footnoteNumFmt = getW(first(first(root, 'footnotePr'), 'numFmt'), 'val');
	const endnoteNumFmt = getW(first(first(root, 'endnotePr'), 'numFmt'), 'val');
	return {
		evenAndOddHeaders: Boolean(first(root, 'evenAndOddHeaders')),
		...(footnoteNumFmt ? { footnoteNumFmt } : {}),
		...(endnoteNumFmt ? { endnoteNumFmt } : {}),
	};
}

async function resolveSlots(
	zip: JSZip,
	relationships: Map<string, Relationship>,
	refs: RawHeaderFooterRef[],
	cache: Map<string, Block[]>,
	fieldMarkup: { seen: boolean },
): Promise<HeaderFooterSlots | undefined> {
	if (!refs.length) return undefined;
	const slots: HeaderFooterSlots = {};
	for (const ref of refs) {
		const relationship = relationships.get(ref.rId);
		if (!relationship) continue;
		const path = resolvePartPath('word/document.xml', relationship.target);
		let blocks = cache.get(path);
		if (!blocks) {
			const xml = await readPart(zip, path);
			if (xml === undefined) continue;
			if (FIELD_MARKUP.test(xml)) fieldMarkup.seen = true;
			blocks = parseBlocksFromContainer(
				parseXml(xml).documentElement,
				`${path.replace(/[^a-z0-9]+/gi, '')}-`,
			);
			cache.set(path, blocks);
		}
		slots[ref.slot] = { blocks };
	}
	return Object.keys(slots).length ? slots : undefined;
}

export interface DocumentPartsResult {
	sections: SectionProperties[];
	footnotes?: Note[];
	endnotes?: Note[];
	footnoteNumFmt?: string;
	endnoteNumFmt?: string;
	evenAndOddHeaders?: boolean;
	warnings: string[];
}

/** Resolves sections (with header/footer content) and footnotes/endnotes for a loaded package. */
export async function parseDocumentParts(
	zip: JSZip,
	body: XmlElement,
	blocks: Block[],
): Promise<DocumentPartsResult> {
	const raw = parseRawSections(body, blocks);
	const relationships = parseRelationships(await readPart(zip, 'word/_rels/document.xml.rels'));
	const settings = parseSettings(await readPart(zip, 'word/settings.xml'));
	const cache = new Map<string, Block[]>();
	const fieldMarkup = { seen: false };
	const sections: SectionProperties[] = [];
	for (const { headerRefs, footerRefs, ...rest } of raw) {
		const headers = await resolveSlots(zip, relationships, headerRefs, cache, fieldMarkup);
		const footers = await resolveSlots(zip, relationships, footerRefs, cache, fieldMarkup);
		sections.push({ ...rest, ...(headers ? { headers } : {}), ...(footers ? { footers } : {}) });
	}
	const footnotesXml = await readPart(zip, 'word/footnotes.xml');
	const endnotesXml = await readPart(zip, 'word/endnotes.xml');
	const footnotes = footnotesXml
		? parseNotesPart(footnotesXml, 'footnote', parseBlocksFromContainer)
		: undefined;
	const endnotes = endnotesXml
		? parseNotesPart(endnotesXml, 'endnote', parseBlocksFromContainer)
		: undefined;

	const warnings: string[] = [];
	if (sections.some((section) => section.headers || section.footers))
		warnings.push(
			'Headers and footers are parsed and displayed read-only; they cannot be edited in this editor.',
		);
	if (sections.some((section) => section.columns.count > 1))
		warnings.push(
			"Section column count, spacing and widths are modeled but rendered as a single continuous column; Word's newspaper-style layout is not reproduced visually.",
		);
	if ((footnotes?.length ?? 0) > 0 || (endnotes?.length ?? 0) > 0)
		warnings.push(
			'Footnote and endnote text is parsed and displayed at the end of the document; paragraphs containing a footnote or endnote reference mark cannot be edited.',
		);
	if (fieldMarkup.seen)
		warnings.push(
			'Field codes such as PAGE and NUMPAGES inside headers/footers are shown as static placeholders and are not recalculated.',
		);
	return {
		sections,
		...(footnotes ? { footnotes } : {}),
		...(endnotes ? { endnotes } : {}),
		...(settings.footnoteNumFmt ? { footnoteNumFmt: settings.footnoteNumFmt } : {}),
		...(settings.endnoteNumFmt ? { endnoteNumFmt: settings.endnoteNumFmt } : {}),
		...(settings.evenAndOddHeaders ? { evenAndOddHeaders: true } : {}),
		warnings,
	};
}
