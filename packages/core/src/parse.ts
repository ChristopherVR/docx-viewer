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
import { parseComments } from './comments.js';
import { parseTrackChangesSetting } from './settings.js';
import { parseRunStyleCatalog } from './character-styles.js';
import { parseTableStyleCatalog } from './table-styles.js';
import { parseTheme, parseColorSchemeMapping } from './theme.js';

const px = (twips: string | undefined, fallback: number): number =>
	twips === undefined ? fallback : (Number(twips) * 96) / 1440;

function hasAny(document: XmlDocument, names: string[]): boolean {
	return names.some((name) => document.getElementsByTagNameNS(WORD_NS, name).length > 0);
}

function warningsFor(document: XmlDocument): string[] {
	const warnings: string[] = [];
	const features: [string[], string][] = [
		[['drawing', 'pict', 'object'], 'Images and drawing objects are preserved but not editable.'],
		[['altChunk'], 'Embedded alternate-format content is not represented in the document model.'],
		[
			['hyperlink'],
			'Hyperlink targets are not represented in the document model; edits inside linked paragraphs are rejected to protect the original XML.',
		],
		[
			['fldSimple', 'instrText', 'fldChar'],
			'Field codes such as PAGE and NUMPAGES are shown as static placeholders (or omitted) and are not recalculated.',
		],
		[
			['moveFrom', 'moveTo'],
			'Moved text is tracked as a paired delete/insert revision; Word’s move linkage between them is not modeled.',
		],
		[
			['rPrChange', 'pPrChange'],
			'Formatting-change revisions are recorded but their prior formatting snapshot is not modeled or rendered; editing the affected run or paragraph drops the recorded snapshot.',
		],
		[
			['tblPrChange', 'trPrChange', 'tcPrChange'],
			'Table-structure tracked changes are preserved in the source XML but are not represented in the document model.',
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
		const stylesXml = await stylesFile.async('string');
		model.paragraphStyles = parseParagraphStyleCatalog(stylesXml);
		model.characterStyles = parseRunStyleCatalog(stylesXml);
		model.tableStyles = parseTableStyleCatalog(stylesXml);
		model.warnings.push(
			'Paragraph style, character style and docDefaults inheritance resolve for rendering, including toggle-property XOR semantics and basedOn chains. Linked styles beyond a basedOn chain, numbering-derived formatting and font metric substitution are not modeled; display can still differ from Word.',
			...model.paragraphStyles.warnings,
			...model.characterStyles.warnings,
			...model.tableStyles.warnings,
		);
	}
	const themeFile = zip.file('word/theme/theme1.xml');
	const settingsFile = zip.file('word/settings.xml');
	if (themeFile) {
		model.theme = parseTheme(await themeFile.async('string'));
		if (settingsFile)
			model.theme.colorMapping = parseColorSchemeMapping(await settingsFile.async('string'));
		model.warnings.push(
			'Theme colors and fonts resolve for rendering through a separate layer; direct theme references are preserved and never flattened onto runs.',
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
			'Table grid widths, merges, borders, shading, cell alignment/margins and table style conditional formatting resolve for rendering; row/column fragmentation and full table-style precedence are not modeled.',
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
	const commentsFile = zip.file('word/comments.xml');
	if (commentsFile) {
		const extendedFile = zip.file('word/commentsExtended.xml');
		model.comments = parseComments(
			await commentsFile.async('string'),
			await extendedFile?.async('string'),
		);
		model.warnings.push(
			'Comments are anchored per paragraph; a comment range spanning multiple paragraphs is not modeled.',
		);
	}
	if (settingsFile) model.trackChanges = parseTrackChangesSetting(await settingsFile.async('string'));
	if (
		blocks.some(
			(block) =>
				block.type === 'table' &&
				block.rows.some((row) => row.some((cell) => cell.nestedTables?.length)),
		)
	)
		model.warnings.push(
			'Nested tables render as a read-only text preview; edit their content from the original document.',
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
