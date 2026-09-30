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
import { parseParagraphStyleCatalog } from './paragraph-styles.js';
import { parseBlocksFromContainer } from './block-parser.js';
import { parseDocumentParts } from './document-parts.js';
import { parseNumberingCatalog } from './numbering-parse.js';
import { parseComments } from './comments.js';
import { parseTrackChangesSetting } from './settings.js';
import { parsePageBackground } from './page-background.js';
import { parseRunStyleCatalog } from './character-styles.js';
import { parseTableStyleCatalog } from './table-styles.js';
import { parseTheme, parseColorSchemeMapping } from './theme.js';

const px = (twips: string | undefined, fallback: number): number =>
	twips === undefined ? fallback : (Number(twips) * 96) / 1440;

import { type DrawingContext } from './drawing.js';
import { parseContentTypes, parseRelationships } from './package-parts.js';
import { withParseWarnings } from './parse-diagnostics.js';
import { warningsFor, imageAndBookmarkWarnings, forEachParagraph } from './parse-warnings.js';

export interface PackageContext {
	original: Uint8Array;
	sourceXml: string;
	base: DocumentModel;
}

export async function readPackage(input: Uint8Array | ArrayBuffer): Promise<{
	model: DocumentModel;
	context: PackageContext;
	media: ReadonlyMap<string, Uint8Array>;
}> {
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
	const relsFile = zip.file('word/_rels/document.xml.rels');
	const contentTypesFile = zip.file('[Content_Types].xml');
	const mediaParts = new Set(Object.keys(zip.files).filter((name) => !zip.files[name]!.dir));
	const drawings: DrawingContext = {
		rels: parseRelationships(relsFile ? await relsFile.async('string') : undefined),
		contentTypes: parseContentTypes(
			contentTypesFile ? await contentTypesFile.async('string') : undefined,
		),
		mediaParts,
	};
	const schemaWarnings: string[] = [];
	const blocks: Block[] = withParseWarnings(schemaWarnings, () =>
		parseBlocksFromContainer(body, '', drawings),
	);
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
		const catalogs = withParseWarnings(schemaWarnings, () => ({
			paragraph: parseParagraphStyleCatalog(stylesXml),
			character: parseRunStyleCatalog(stylesXml),
			table: parseTableStyleCatalog(stylesXml),
		}));
		model.paragraphStyles = catalogs.paragraph;
		model.characterStyles = catalogs.character;
		model.tableStyles = catalogs.table;
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
		if (settingsFile) {
			const settingsXml = await settingsFile.async('string');
			model.theme.colorMapping = withParseWarnings(schemaWarnings, () =>
				parseColorSchemeMapping(settingsXml),
			);
		}
		model.warnings.push(
			'Theme colors and fonts resolve for rendering through a separate layer; direct theme references are preserved and never flattened onto runs.',
		);
	}
	const numberingFile = zip.file('word/numbering.xml');
	if (numberingFile) {
		const numberingXml = await numberingFile.async('string');
		model.numberingCatalog = withParseWarnings(schemaWarnings, () =>
			parseNumberingCatalog(numberingXml),
		);
		model.warnings.push(
			'Numbering is resolved for decimal, Roman numeral, letter, ordinal, cardinal/ordinal text and bullet formats, including multilevel lvlText and legal numbering. Picture bullets, style-linked numbering and other custom formats fall back to Decimal Number and are not rendered as Word would.',
			...model.numberingCatalog.warnings,
		);
	} else if (document.getElementsByTagNameNS(WORD_NS, 'numPr').length > 0) {
		model.warnings.push(
			'Paragraphs reference list numbering, but the package has no word/numbering.xml part; numbering is preserved as paragraph XML but not rendered.',
		);
	}
	if (blocks.some((block) => block.type === 'table'))
		model.warnings.push(
			'Table grid widths, merges, borders, shading, cell alignment/margins and table style conditional formatting resolve for rendering; row/column fragmentation and full table-style precedence are not modeled.',
		);
	const parts = await parseDocumentParts(zip, body, blocks, drawings);
	model.sections = parts.sections;
	if (parts.footnotes) model.footnotes = parts.footnotes;
	if (parts.endnotes) model.endnotes = parts.endnotes;
	if (parts.footnoteNumFmt) model.footnoteNumFmt = parts.footnoteNumFmt;
	if (parts.endnoteNumFmt) model.endnoteNumFmt = parts.endnoteNumFmt;
	if (parts.evenAndOddHeaders) model.evenAndOddHeaders = true;
	if (parts.autoHyphenation) model.autoHyphenation = true;
	const pageColor = parsePageBackground(document.documentElement);
	if (pageColor) model.pageColor = pageColor;
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
			'Comment ranges are modeled at run granularity (they may span paragraphs); a range edge outside any run, such as around an empty paragraph or a table, moves to the nearest commented text when that paragraph is edited.',
		);
	}
	if (settingsFile)
		model.trackChanges = parseTrackChangesSetting(await settingsFile.async('string'));
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
	model.warnings.push(...imageAndBookmarkWarnings(blocks));
	for (const warning of schemaWarnings)
		if (!model.warnings.includes(warning)) model.warnings.push(warning);
	// Pictures anywhere in the document: body, headers, footers, footnotes and endnotes.
	const imagePartNames = new Set<string>();
	const pictureBlocks = [
		...blocks,
		...(model.sections ?? []).flatMap((section) =>
			[section.headers, section.footers].flatMap((slots) =>
				Object.values(slots ?? {}).flatMap((content) => content?.blocks ?? []),
			),
		),
		...[...(model.footnotes ?? []), ...(model.endnotes ?? [])].flatMap((note) => note.blocks),
	];
	forEachParagraph(pictureBlocks, (paragraph) => {
		for (const run of paragraph.runs)
			for (const part of [run.image?.partName, run.image?.svgPartName])
				if (part) imagePartNames.add(part);
	});
	const media = new Map<string, Uint8Array>();
	for (const partName of imagePartNames) {
		const part = zip.file(partName);
		if (part) media.set(partName, await part.async('uint8array'));
	}
	const context = { original, sourceXml, base: structuredClone(model) };
	remember(model, context);
	return { model, context, media };
}

export async function loadDocx(input: Uint8Array | ArrayBuffer): Promise<LoadedDocument> {
	const { model, context, media } = await readPackage(input);
	return {
		model,
		media,
		save: (next = model, pendingMedia) => {
			if (next !== model) remember(next, context);
			return saveDocx(next, pendingMedia);
		},
	};
}
