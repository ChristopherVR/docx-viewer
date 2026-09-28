// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Orchestrates section, header/footer and footnote/endnote parsing on top of block-parser.ts,
// sections.ts and notes.ts. Kept out of parse.ts so the hub file only needs one hook call.
import type JSZip from 'jszip';
import type { Block, HeaderFooterSlots, Note, SectionProperties } from './model.js';
import { first, getW, parseXml, type XmlElement } from './xml.js';
import { isStNumberFormat, type StNumberFormat } from './generated/wml-simple-types.js';
import { enumValue, withParseWarnings } from './parse-diagnostics.js';
import { parseRelationships, resolvePartPath, type Relationship } from './relationships.js';
import { parseRawSections, type RawHeaderFooterRef } from './sections.js';
import { parseNotesPart } from './notes.js';
import { parseBlocksFromContainer } from './block-parser.js';
import type { DrawingContext } from './drawing.js';
import { parseRelationships as parsePackageRelationships } from './package-parts.js';

/** A part's own relationships (word/_rels/<part>.rels) for resolving its pictures and links. */
async function partContext(
	zip: JSZip,
	path: string,
	base: DrawingContext | undefined,
): Promise<DrawingContext | undefined> {
	if (!base) return undefined;
	const slash = path.lastIndexOf('/');
	const relsPath = `${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`;
	return { ...base, rels: parsePackageRelationships(await readPart(zip, relsPath)) };
}

const FIELD_MARKUP = /<w:(?:fldSimple|instrText|fldChar)\b/;

async function readPart(zip: JSZip, path: string): Promise<string | undefined> {
	return zip.file(path)?.async('string');
}

interface SettingsInfo {
	evenAndOddHeaders: boolean;
	footnoteNumFmt?: StNumberFormat;
	endnoteNumFmt?: StNumberFormat;
}
function parseSettings(xml: string | undefined): SettingsInfo {
	if (!xml) return { evenAndOddHeaders: false };
	const root = parseXml(xml).documentElement;
	const footnoteNumFmt = enumValue(
		isStNumberFormat,
		getW(first(first(root, 'footnotePr'), 'numFmt'), 'val'),
		'w:footnotePr/w:numFmt',
	);
	const endnoteNumFmt = enumValue(
		isStNumberFormat,
		getW(first(first(root, 'endnotePr'), 'numFmt'), 'val'),
		'w:endnotePr/w:numFmt',
	);
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
	drawings: DrawingContext | undefined,
	sink: string[],
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
			const context = await partContext(zip, path, drawings);
			const root = parseXml(xml).documentElement;
			blocks = withParseWarnings(sink, () =>
				parseBlocksFromContainer(root, `${path.replace(/[^a-z0-9]+/gi, '')}-`, context),
			);
			cache.set(path, blocks);
		}
		slots[ref.slot] = { blocks, partName: path };
	}
	return Object.keys(slots).length ? slots : undefined;
}

export interface DocumentPartsResult {
	sections: SectionProperties[];
	footnotes?: Note[];
	endnotes?: Note[];
	footnoteNumFmt?: StNumberFormat;
	endnoteNumFmt?: StNumberFormat;
	evenAndOddHeaders?: boolean;
	warnings: string[];
}

/** Resolves sections (with header/footer content) and footnotes/endnotes for a loaded package. */
export async function parseDocumentParts(
	zip: JSZip,
	body: XmlElement,
	blocks: Block[],
	drawings?: DrawingContext,
): Promise<DocumentPartsResult> {
	const warnings: string[] = [];
	const raw = withParseWarnings(warnings, () => parseRawSections(body, blocks));
	const relationships = parseRelationships(await readPart(zip, 'word/_rels/document.xml.rels'));
	const settingsXml = await readPart(zip, 'word/settings.xml');
	const settings = withParseWarnings(warnings, () => parseSettings(settingsXml));
	const cache = new Map<string, Block[]>();
	const fieldMarkup = { seen: false };
	const sections: SectionProperties[] = [];
	for (const { headerRefs, footerRefs, ...rest } of raw) {
		const headers = await resolveSlots(
			zip,
			relationships,
			headerRefs,
			cache,
			fieldMarkup,
			drawings,
			warnings,
		);
		const footers = await resolveSlots(
			zip,
			relationships,
			footerRefs,
			cache,
			fieldMarkup,
			drawings,
			warnings,
		);
		sections.push({ ...rest, ...(headers ? { headers } : {}), ...(footers ? { footers } : {}) });
	}
	const footnotesXml = await readPart(zip, 'word/footnotes.xml');
	const endnotesXml = await readPart(zip, 'word/endnotes.xml');
	const notesParser = async (path: string) => {
		const context = await partContext(zip, path, drawings);
		return (container: XmlElement, prefix: string) =>
			withParseWarnings(warnings, () => parseBlocksFromContainer(container, prefix, context));
	};
	const footnotes = footnotesXml
		? parseNotesPart(footnotesXml, 'footnote', await notesParser('word/footnotes.xml'))
		: undefined;
	const endnotes = endnotesXml
		? parseNotesPart(endnotesXml, 'endnote', await notesParser('word/endnotes.xml'))
		: undefined;

	if (sections.some((section) => section.columns.count > 1))
		warnings.push(
			'Columns render in Print Layout and, for single-section documents, on the editing surface; column breaks balance approximately.',
		);
	if (fieldMarkup.seen)
		warnings.push(
			'PAGE, NUMPAGES, SECTIONPAGES, DATE and TIME fields are recalculated in Print Layout; other fields show the result Word last saved. Field results are editable and field codes are preserved.',
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
