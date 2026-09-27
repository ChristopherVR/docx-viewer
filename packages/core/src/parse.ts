// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import JSZip from 'jszip';
import type { Block, DocumentModel, LoadedDocument } from './model.js';
import {
	children,
	first,
	getW,
	parseXml,
	type XmlDocument,
	type XmlElement,
	WORD_NS,
} from './xml.js';
import { remember, saveDocx } from './save.js';
import { hasSpecialBreak, isModeledBreak } from './breaks.js';
import { parseParagraphStyleCatalog } from './paragraph-styles.js';
import { parseBlocksFromContainer } from './block-parser.js';
import { parseDocumentParts } from './document-parts.js';
import { parseNumberingCatalog } from './numbering-parse.js';

const px = (twips: string | undefined, fallback: number): number =>
	twips === undefined ? fallback : (Number(twips) * 96) / 1440;

function hasAny(document: XmlDocument, names: string[]): boolean {
	return names.some((name) => document.getElementsByTagNameNS(WORD_NS, name).length > 0);
}

function warningsFor(document: XmlDocument): string[] {
	const warnings: string[] = [];
	const features: [string[], string][] = [
		[['drawing', 'pict', 'object'], 'Images and drawing objects are preserved but not editable.'],
		[
			['commentRangeStart', 'trackRevisions'],
			'Comments and tracked review features are not represented in the document model.',
		],
		[['altChunk'], 'Embedded alternate-format content is not represented in the document model.'],
		[
			['hyperlink'],
			'Hyperlink targets are not represented in the document model; edits inside linked paragraphs are rejected to protect the original XML.',
		],
		[
			['rStyle'],
			'Character style inheritance and theme font/color resolution are not modeled; displayed formatting may differ from Word.',
		],
		[
			['fldSimple', 'instrText', 'fldChar'],
			'Field codes such as PAGE and NUMPAGES are shown as static placeholders (or omitted) and are not recalculated.',
		],
	];
	for (const [names, message] of features) if (hasAny(document, names)) warnings.push(message);
	const specialBreak =
		Array.from(document.getElementsByTagNameNS(WORD_NS, 'br')).some(
			(br: XmlElement) => hasSpecialBreak(br) && !isModeledBreak(br),
		) ||
		Array.from(document.getElementsByTagNameNS(WORD_NS, 'cr')).some(
			(cr: XmlElement) => cr.attributes.length > 0,
		);
	if (specialBreak)
		warnings.push(
			'Some non-line breaks other than page and column breaks are not distinguished from line breaks in the document model; edits to paragraphs containing them are rejected to preserve the original XML.',
		);
	return warnings;
}

export interface PackageContext {
	original: Uint8Array;
	sourceXml: string;
	base: DocumentModel;
}

export async function readPackage(
	input: Uint8Array | ArrayBuffer,
): Promise<{ model: DocumentModel; context: PackageContext }> {
	const original =
		input instanceof Uint8Array ? new Uint8Array(input) : new Uint8Array(input.slice(0));
	if (original.byteLength > 50 * 1024 * 1024)
		throw new Error('DOCX package exceeds the 50 MiB compressed input limit');
	const zip = await JSZip.loadAsync(original);
	const entries = Object.values(zip.files);
	if (entries.length > 5000) throw new Error('DOCX package exceeds the 5,000 part limit');
	const totalSize = entries.reduce(
		(sum, entry) => sum + Number((entry as any)._data?.uncompressedSize ?? 0),
		0,
	);
	if (totalSize > 200 * 1024 * 1024)
		throw new Error('DOCX package exceeds the 200 MiB uncompressed content limit');
	const file = zip.file('word/document.xml');
	if (!file) throw new Error('DOCX package has no word/document.xml part');
	const sourceXml = await file.async('string');
	if (sourceXml.length > 25 * 1024 * 1024)
		throw new Error('word/document.xml exceeds the 25 MiB limit');
	const document: XmlDocument = parseXml(sourceXml);
	const body = Array.from(document.getElementsByTagNameNS(WORD_NS, 'body'))[0];
	if (!body) throw new Error('DOCX document.xml has no w:body');
	const blocks: Block[] = parseBlocksFromContainer(body);
	const section = children(body, 'sectPr').at(-1);
	const size = first(section, 'pgSz');
	const margins = first(section, 'pgMar');
	const model: DocumentModel = {
		blocks,
		page: {
			width: px(getW(size, 'w'), 816),
			height: px(getW(size, 'h'), 1056),
			marginTop: px(getW(margins, 'top'), 96),
			marginRight: px(getW(margins, 'right'), 96),
			marginBottom: px(getW(margins, 'bottom'), 96),
			marginLeft: px(getW(margins, 'left'), 96),
		},
		warnings: warningsFor(document),
	};
	const stylesFile = zip.file('word/styles.xml');
	if (stylesFile) {
		model.paragraphStyles = parseParagraphStyleCatalog(await stylesFile.async('string'));
		model.warnings.push(
			'Paragraph style inheritance is resolved for alignment, direction, spacing and indentation. Run formatting, character styles and theme values remain unresolved; display can differ from Word.',
			...model.paragraphStyles.warnings,
		);
	}
	const numberingFile = zip.file('word/numbering.xml');
	if (numberingFile) {
		model.numberingCatalog = parseNumberingCatalog(await numberingFile.async('string'));
		model.warnings.push(
			'Numbering is resolved for decimal, Roman numeral, letter, ordinal, cardinal/ordinal text and bullet formats, including multilevel lvlText and legal numbering. Picture bullets, style-linked numbering and other custom formats fall back to Decimal Number and are not rendered as Word would.',
			...model.numberingCatalog.warnings,
		);
	} else if (hasAny(document, ['numPr'])) {
		model.warnings.push(
			'Paragraphs reference list numbering, but the package has no word/numbering.xml part; numbering is preserved as paragraph XML but not rendered.',
		);
	}
	if (blocks.some((block) => block.type === 'table'))
		model.warnings.push(
			'Table text and cell structure are supported; table widths, borders, shading and cell formatting are not modeled.',
		);
	const parts = await parseDocumentParts(zip, body, blocks);
	model.sections = parts.sections;
	if (parts.footnotes) model.footnotes = parts.footnotes;
	if (parts.endnotes) model.endnotes = parts.endnotes;
	if (parts.footnoteNumFmt) model.footnoteNumFmt = parts.footnoteNumFmt;
	if (parts.endnoteNumFmt) model.endnoteNumFmt = parts.endnoteNumFmt;
	if (parts.evenAndOddHeaders) model.evenAndOddHeaders = true;
	model.warnings.push(...parts.warnings);
	if (blocks.some((block) => block.type === 'table' && !block.structureEditable))
		model.warnings.push(
			'Merged, nested, or complex tables can be read, but their row and column structure cannot be edited safely.',
		);
	const context = { original, sourceXml, base: structuredClone(model) };
	remember(model, context);
	return { model, context };
}

export async function loadDocx(input: Uint8Array | ArrayBuffer): Promise<LoadedDocument> {
	const { model, context } = await readPackage(input);
	return {
		model,
		save: (next = model) => {
			if (next !== model) remember(next, context);
			return saveDocx(next);
		},
	};
}
